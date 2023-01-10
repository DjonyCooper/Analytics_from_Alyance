from PyQt5.QtWidgets import (QLineEdit, QPushButton, QLabel, QWidget, QApplication, QFrame,
                             QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout, QStyle)
from PyQt5.QtGui import QFont, QPixmap, QIcon
from PyQt5.QtCore import QDateTime, Qt
import openpyxl
from openpyxl import load_workbook
import time
import datetime

class Main(QWidget):
    def __init__(self):
        super(Main, self).__init__()
        self.setWindowTitle("Тестовая программа: Аналитика v1.0")
        self.setMinimumSize(700, 300)
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_FileDialogStart))
        vbox = QVBoxLayout()
        hbox_line_1 = QHBoxLayout()
        hbox_line_2 = QHBoxLayout()
        hbox_line_3 = QHBoxLayout()
        hbox_line_4 = QHBoxLayout()
        hbox_line_5 = QHBoxLayout()
        hbox_ok_close = QHBoxLayout()
        vbox.addWidget(self.hello_label())
        vbox.addStretch(0)
        vbox.addWidget(self.horizontal_line())
        hbox_line_1.addWidget(self.book_label())
        hbox_line_1.addWidget(self.book_line_edit())
        vbox.addLayout(hbox_line_1)
        hbox_line_2.addWidget(self.book_ost_label())
        hbox_line_2.addWidget(self.book_ost_line_edit())
        vbox.addLayout(hbox_line_2)
        hbox_line_3.addWidget(self.book_price_label())
        hbox_line_3.addWidget(self.book_price_line_edit())
        vbox.addLayout(hbox_line_3)
        hbox_line_4.addWidget(self.book_const_label())
        hbox_line_4.addWidget(self.book_const_line_edit())
        vbox.addLayout(hbox_line_4)
        vbox.addWidget(self.horizontal_line())
        hbox_line_5.addWidget(self.book_file_save_label())
        hbox_line_5.addWidget(self.book_file_save_line_edit())
        vbox.addLayout(hbox_line_5)
        vbox.setSpacing(10)
        vbox.addWidget(self.horizontal_line())
        hbox_ok_close.addWidget(self.b_create())
        hbox_ok_close.addWidget(self.b_close())
        vbox.addLayout(hbox_ok_close)

        self.setLayout(vbox)

    def horizontal_line(self):
        self.decor_line = QFrame()
        self.decor_line.setFrameShape(QFrame.HLine)
        self.decor_line.setFrameShadow(QFrame.Sunken)
        return self.decor_line

    def hello_label(self):
        hello_label = QLabel("Аналитика v1.0")
        hello_label.setFont(QFont('Century Gothic', 20))
        hello_label.setAlignment(Qt.AlignCenter)
        return hello_label

    def book_label(self):
        book_label = QLabel("Выберите файл, содержащий остатки из 1С:")
        book_label.setMinimumSize(300, 10)
        return book_label

    def book_line_edit(self):
        self.book_line_edit = QLineEdit(self)
        self.book_line_edit.setMinimumSize(30, 30)
        serch_book_icon = self.book_line_edit.addAction(self.style().standardIcon(QStyle.SP_FileDialogContentsView), QLineEdit.TrailingPosition)
        serch_book_icon.triggered.connect(self.browse_files_book)
        return self.book_line_edit

    def book_ost_label(self):
        book_ost_label = QLabel("Выберите файл, содержащий остатки из базы:")
        book_ost_label.setMinimumSize(300, 10)
        return book_ost_label

    def book_ost_line_edit(self):
        self.book_ost_line_edit = QLineEdit(self)
        self.book_ost_line_edit.setMinimumSize(30, 30)
        serch_book_ost_icon = self.book_ost_line_edit.addAction(self.style().standardIcon(QStyle.SP_FileDialogContentsView), QLineEdit.TrailingPosition)
        serch_book_ost_icon.triggered.connect(self.browse_files_book_ost)
        return self.book_ost_line_edit

    def book_price_label(self):
        book_price = QLabel("Выберите файл, содержащий прайс:")
        book_price.setMinimumSize(300, 10)
        return book_price

    def book_price_line_edit(self):
        self.book_price_line_edit = QLineEdit(self)
        self.book_price_line_edit.setMinimumSize(30, 30)
        serch_book_price_icon = self.book_price_line_edit.addAction(self.style().standardIcon(QStyle.SP_FileDialogContentsView), QLineEdit.TrailingPosition)
        serch_book_price_icon.triggered.connect(self.browse_files_book_price)
        return self.book_price_line_edit

    def book_const_label(self):
        book_const = QLabel("Выберите файл, содержащий константу:")
        book_const.setMinimumSize(300, 10)
        return book_const

    def book_const_line_edit(self):
        self.book_const_line_edit = QLineEdit(self)
        self.book_const_line_edit.setMinimumSize(30, 30)
        serch_book_const_icon = self.book_const_line_edit.addAction(self.style().standardIcon(QStyle.SP_FileDialogContentsView), QLineEdit.TrailingPosition)
        serch_book_const_icon.triggered.connect(self.browse_files_book_const)
        return self.book_const_line_edit

    def book_file_save_label(self):
        book_file_save_label = QLabel("Выберите папку, для сохранения новых файлов:")
        book_file_save_label.setMinimumSize(300, 10)
        return book_file_save_label

    def book_file_save_line_edit(self):
        self.book_file_save_line_edit = QLineEdit(self)
        self.book_file_save_line_edit.setMinimumSize(30, 30)
        serch_book_ost_icon = self.book_file_save_line_edit.addAction(self.style().standardIcon(QStyle.SP_DialogSaveButton), QLineEdit.TrailingPosition)
        serch_book_ost_icon.triggered.connect(self.browse_files_book_file_save)
        return self.book_file_save_line_edit

    def browse_files_book(self):
        browse_files_book = QFileDialog.getOpenFileName(self, 'Выберите файл, содержащий остатки из 1С...', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_line_edit.setText(browse_files_book[0])

    def browse_files_book_ost(self):
        browse_files_book_ost = QFileDialog.getOpenFileName(self, 'Выберите файл, содержащий остатки из базы...', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_ost_line_edit.setText(browse_files_book_ost[0])

    def browse_files_book_price(self):
        browse_files_book_price = QFileDialog.getOpenFileName(self, 'Выберите файл, содержащий прайс...', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_price_line_edit.setText(browse_files_book_price[0])

    def browse_files_book_const(self):
        browse_files_book_const = QFileDialog.getOpenFileName(self, 'Выберите файл, содержащий константу... ', 'C:\Program Files', 'xlsx files (*.xlsx)')
        self.book_const_line_edit.setText(browse_files_book_const[0])

    def browse_files_book_file_save(self):
        browse_files_book_file_save = QFileDialog.getExistingDirectory(self, "Выбор папки для сохранения...")
        self.book_file_save_line_edit.setText(browse_files_book_file_save)

    def b_create(self):
        b_create = QPushButton("Выполнить обработку • Enter", self)
        b_create.setMinimumSize(10, 40)
        b_create.setFont(QFont('Century Gothic', 8, QFont.Normal))
        b_create.setToolTip('Нажмите, для запуска программы')
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
        b_close.setMinimumSize(50, 40)
        b_close.setFont(QFont('Century Gothic', 8, QFont.Normal))
        b_close.setShortcut('Esc')
        b_close.setToolTip('Нажмите, чтобы выйти')
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
        self.base_line_edit = [self.book_line_edit, self.book_price_line_edit, self.book_const_line_edit, self.book_ost_line_edit, self.book_file_save_line_edit]
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
        msgBox.setWindowIcon(self.style().standardIcon(QStyle.SP_FileDialogStart))
        msgBox.setIcon(QMessageBox.Warning)
        msgBox.setDetailedText('Для корректной работы программы необходимо заполнить все поля, после чего нажать на кнопку "Выполнить обработку"')
        msgBox.setText(message)
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec_()

    def image_and(self):
        label = QLabel(self)
        pixmap = QPixmap('line.png')
        label.setPixmap(pixmap)
        return label

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

        book_rez.save(f'{self.book_file_save_line_edit.text()}/44_Остатки(1) {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_rez.save(f'{self.book_file_save_line_edit.text()}/ost_base.xlsx')
        book_al.save(f'{self.book_file_save_line_edit.text()}/46учетАльянс_Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_44_dealer.save(f'{self.book_file_save_line_edit.text()}/44дилер_Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
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

        book_46rez.save(f'{self.book_file_save_line_edit.text()}/44_Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_46_bosch_rez.save(f'{self.book_file_save_line_edit.text()}/46_BOSCH.xlsx')
        book_46_dealer_rez.save(f'{self.book_file_save_line_edit.text()}/46_дилер_Остатки {datetime.date.today().strftime("%d.%m.%y")}.xlsx')
        book_44_bosch_rez.save(f'{self.book_file_save_line_edit.text()}/44_BOSCH.xlsx')
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
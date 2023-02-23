if __name__ == '__main__':
    import os
    import pandas as pd
    from tika import parser
    import comtypes.client

    from kivy.uix.popup import Popup
    from kivy.uix.label import Label
    from kivy.core.window import Window
    from kivy.lang.builder import Builder

    from kivymd.app import MDApp
    from kivymd.toast import toast
    from kivymd.uix.tooltip import MDTooltip
    from kivymd.uix.button import MDIconButton
    from kivymd.uix.filemanager import MDFileManager

    Window.minimum_height = 250
    Window.minimum_width = 400
    Window.size = (500, 500)


    class TooltipMDIconButton(MDIconButton, MDTooltip):
        pass


    class MainApp(MDApp):
        def __init__(self, **kwargs):
            super().__init__(**kwargs)
            Window.bind(on_keyboard=self.events)
            self.app = None
            self.path_report = None
            self.path_doc_files = None
            self.path_pdf_files = None

            self.manager_report_open = False
            self.file_manager_report = MDFileManager(
                exit_manager=self.exit_manager_report,
                select_path=self.select_path_report
            )

            self.manager_converter_open_doc = False
            self.file_manager_converter_doc = MDFileManager(
                exit_manager=self.exit_manager_converter_doc,
                select_path=self.select_path_converter_doc
            )

            self.manager_converter_open_pdf = False
            self.file_manager_converter_pdf = MDFileManager(
                exit_manager=self.exit_manager_converter_pdf,
                select_path=self.select_path_converter_pdf
            )

            self.popup_warning = Popup(
                title='Предупреждение',
                title_size=20,
                content=Label(text='', font_size=20),
                size_hint=(None, None),
                size=(350, 150),
                auto_dismiss=True
            )

            self.popup_message = Popup(
                title='Уведомление',
                title_size=20,
                content=Label(text='', font_size=20),
                size_hint=(None, None),
                size=(350, 150),
                auto_dismiss=True
            )

        def build(self):
            self.app = Builder.load_file("kv_main.kv")
            self.title = ''
            return self.app

        def file_manager_report_open(self):
            self.file_manager_report.show('/')
            self.manager_report_open = True

        def select_path_report(self, path):
            self.exit_manager_report()
            toast(path)
            self.path_report = os.path.abspath(path)
            self.app.ids.report_label.text = self.path_report

        def exit_manager_report(self, *args):
            self.manager_report_open = False
            self.file_manager_report.close()

        def file_manager_converter_open_doc(self):
            self.file_manager_converter_doc.show('/')
            self.manager_converter_open_doc = True

        def select_path_converter_doc(self, path):
            self.exit_manager_converter_doc()
            toast(path)
            self.path_doc_files = os.path.abspath(path)
            self.app.ids.converter_label_doc.text = f'DOC - {self.path_doc_files}'

        def exit_manager_converter_doc(self, *args):
            self.manager_converter_open_doc = False
            self.file_manager_converter_doc.close()

        def file_manager_converter_open_pdf(self):
            self.file_manager_converter_pdf.show('/')
            self.manager_converter_open_pdf = True

        def select_path_converter_pdf(self, path):
            self.exit_manager_converter_pdf()
            toast(path)
            self.path_pdf_files = os.path.abspath(path)
            self.app.ids.converter_label_pdf.text = f'PDF - {self.path_pdf_files}\\Файлы pdf'

        def exit_manager_converter_pdf(self, *args):
            self.manager_converter_open_pdf = False
            self.file_manager_converter_pdf.close()

        def events(self, instance, keyboard, keycode, text, modifiers):
            if keyboard in (1001, 27):
                if self.manager_converter_open_doc:
                    self.file_manager_converter_doc.back()

                if self.manager_converter_open_pdf:
                    self.file_manager_converter_pdf.back()

                if self.manager_report_open:
                    self.file_manager_report.back()
            return True

        def create_report(self):

            if not self.path_report:
                self.popup_warning.content.text = 'Выберите директорию отчета'
                self.popup_warning.open()
                return

            if not self.path_pdf_files:
                self.popup_warning.content.text = 'Выберите директорию pdf файлов'
                self.popup_warning.open()
                return

            words = {
                'Добыча': 'добыч',
                'Транспортировка': 'транспортировк',
                'Переработка': 'переработк',
                'Хранение': 'хранен'
            }
            marks = '''!()-[]{};?@#$%:'"\,./^&amp;*_'''
            list_date_collapse = []
            list_name_organisation = []
            list_reason_collapse = []
            list_org_reason = []
            list_deaths = []
            list_travm = []
            list_money = []

            for year in os.listdir(self.path_pdf_files):
                for file in os.listdir(f'{self.path_pdf_files}\\{year}'):
                    money = 0
                    raw = parser.from_file(f'{self.path_pdf_files}\\{year}\\{file}')
                    content = raw['content']  # хранит pdf файл

                    if 'Дата происшествия:' in content and '\nНаименование \nорганизации:' not in content:
                        date_collapse = content[content.find('Дата происшествия:') + len('Дата происшествия:') + 1:]
                        date_collapse = date_collapse[:date_collapse.find('\n')]
                    elif 'Дата происшествия:' in content:
                        date_collapse = content[
                                        content.find('Дата происшествия:') + len('Дата происшествия: '):content.find(
                                            '\nНаименование \nорганизации:')]
                        date_collapse = date_collapse.rstrip('\n')
                    elif 'Дата и время \nпроисшествия:' in content:
                        date_collapse = content[content.find('Дата и время \nпроисшествия:') + len(
                            "Дата и время \nпроисшествия:") + 2:content.find('\n\nНаименование \nорганизации:')]
                        date_collapse = date_collapse[:date_collapse.find(' ')]
                    elif 'Дата \nпроисшествия:' in content and '\nНаименование \nорганизации:' in content:
                        date_collapse = content[
                                        content.find('Дата \nпроисшествия:') + len(
                                            'Дата \nпроисшествия:') + 2:content.find(
                                            '\nНаименование \nорганизации:') - 1]
                    elif 'Дата происшествия:' not in content and 'Дата и время' not in content and 'Дата \nпроисшествия:' not in content:
                        date_collapse = content[
                                        content.find('У Р О К И ,  И З В Л Е Ч Е Н Н Ы Е  И З  А В А Р И И') + len(
                                            'У Р О К И ,  И З В Л Е Ч Е Н Н Ы Е  И З  А В А Р И И') + 2:]
                        date_collapse = date_collapse[:date_collapse.find('\n')]
                    else:
                        date_collapse = content[content.find('Дата \nпроисшествия:') + len('Дата происшествия: ') + 1:]
                        date_collapse = date_collapse[:date_collapse.find('\n')]

                    if 'Наименование \nорганизации:' in content:
                        name_organisation = content[content.find('Наименование \nорганизации:') + len(
                            'Наименование \nорганизации:'):content.find('Ведомственная \nпринадлежность')]
                    elif 'Наименование\nорганизации:' in content:
                        name_organisation = content[content.find('Наименование\nорганизации:') + len(
                            'Наименование\nорганизации:'):content.find('Ведомственная\nпринадлежность')]

                    elif 'Наименование \nорганизации:' not in content:
                        name_organisation = content[content.find(date_collapse) + len(date_collapse):content.find(
                            'Ведомственная \nпринадлежность')]

                    name_organisation = name_organisation.replace('\n', '')
                    name_organisation = name_organisation.strip(' ')

                    f = 1
                    for i, j in words.items():
                        if j in content:
                            reason_collapse = i
                            f = 0
                            break
                    if f:
                        reason_collapse = 'Иное'
                        f = 1
                    if 'газораспределение' in file.lower():
                        reason_collapse = 'Транспортировка'

                    list_date_collapse.append(date_collapse)
                    list_name_organisation.append(name_organisation)
                    list_reason_collapse.append(reason_collapse)
                    death = 0
                    travm = 0
                    if 'Пострадавших нет' in content:
                        1
                    else:
                        view = content[
                               content.find('Последствия аварии') + len('Последствия аварии'):content.find(
                                   'Технические причины')]
                        view_sp = view.split()

                        for mark in marks:
                            for j in range(len(view_sp)):
                                view_sp[j] = view_sp[j].replace(mark, '')
                        if 'Погиб' in view_sp or 'погиб' in view_sp:
                            death += 1

                        if 'Погибли' in view_sp:
                            death += sum(
                                [int(s) for s in view[view.find('Погибли') + len('Погибли'):view.find('Погибли') + len(
                                    'Погибли') + 30].split() if s.isdigit()])
                        elif 'погибли' in view_sp:
                            death += sum(
                                [int(s) for s in view[view.find('погибли') + len('погибли'):view.find('погибли') + len(
                                    'погибли') + 30].split() if s.isdigit()])

                        if 'острад' in view:
                            travm += sum([int(s) for s in
                                          view[view.find('острад') + len('острад'):view.find('острад') + len(
                                              'острад') + 40].split()
                                          if s.isdigit()])
                        if 'пострадало' in view:
                            travm += sum([int(s) for s in view[
                                                          view.find('пострадало') + len('пострадало'):view.find(
                                                              'пострадало') + len(
                                                              'пострадало') + 40].split() if s.isdigit()])
                        if 'пострадали' in view:
                            travm += sum([int(s) for s in view[
                                                          view.find('пострадали') + len('пострадали'):view.find(
                                                              'пострадали') + len(
                                                              'пострадали') + 20].split() if s.isdigit()])

                        if 'Пострадали трое' in view:
                            travm += 3
                        if 'двое - \nсмертельно' in view:
                            death += 2
                        if 'Смертельно травмированы четверо' in view:
                            death += 4
                        if 'Пострадало 4' in view:
                            travm += 4
                        if 'Пострадал оператор' in view:
                            travm += 4
                        if 'Двое пострадавших' in view:
                            travm += 2
                        if 'получили 4 \nработника' in view:
                            travm += 4
                        if 'Машинист' in view:
                            travm += 1

                        if '3 - смертельно' in view:
                            death += 3
                        if 'скончался' in view:
                            death += 1
                        if '1 человек - смертельно' in view:
                            death += 1

                        travm += view.count('травмирован') + view.count('Травмирован')
                    list_deaths.append(death)
                    list_travm.append(travm)

                    view = content[
                           content.find('Последствия аварии') + len('Последствия аварии'):content.find(
                               'Технические причины')]
                    money = 0
                    f_k = 0
                    f_kk = 0
                    kof = 1
                    if view.count('руб.') == 1:
                        stroka = view[view.find('руб.') - 20:view.find('руб.') + 4].split()
                        for j in range(len(stroka)):
                            stroka[j] = stroka[j].replace(',', '.')
                            stroka[j] = stroka[j].replace(' ', '')
                        if 'тыс.' in stroka or 'тыс.руб.' in stroka or 'тыс' in stroka:
                            f_k = 1
                        if 'млн.' in stroka or 'млн' in stroka:
                            f_kk = 1
                        stroka_num = ''
                        for j in range(len(stroka)):
                            try:
                                float(stroka[j])
                                stroka_num += stroka[j]
                            except:
                                continue
                        money += float(stroka_num)
                        if f_k:
                            kof = 1000
                        elif f_kk:
                            kof = 1000000
                        money *= kof
                    elif view.count('рублей') == 1:
                        stroka = view[view.find('рублей') - 20:view.find('рублей') + 6].split()
                        for j in range(len(stroka)):
                            stroka[j] = stroka[j].replace(',', '.')
                            stroka[j] = stroka[j].replace(' ', '')
                        if 'тыс.' in stroka or 'тыс.рублей' in stroka or 'тыс' in stroka:
                            f_k = 1
                        if 'млн.' in stroka or 'млн' in stroka:
                            f_kk = 1
                        stroka_num = ''
                        for j in range(len(stroka)):
                            try:
                                float(stroka[j])
                                stroka_num += stroka[j]
                            except:
                                continue
                        money += float(stroka_num)
                        if f_k:
                            kof = 1000
                        elif f_kk:
                            kof = 1000000
                        money *= kof

                    elif view.count('р.') == 1:
                        stroka = view[view.find('р.') - 20:view.find('р.') + 2].split()
                        for j in range(len(stroka)):
                            stroka[j] = stroka[j].replace(',', '.')
                            stroka[j] = stroka[j].replace(' ', '')
                            stroka[j] = stroka[j].replace('р.', '')
                        if 'тыс.' in stroka or 'тыс.р.' in stroka or 'тыс' in stroka:
                            f_k = 1
                        if 'млн.' in stroka or 'млн' in stroka:
                            f_kk = 1
                        stroka_num = ''
                        for j in range(len(stroka)):
                            try:
                                float(stroka[j])
                                stroka_num += stroka[j]
                            except:
                                continue
                        money += float(stroka_num)
                        if f_k:
                            kof = 1000
                        elif f_kk:
                            kof = 1000000
                        money *= kof

                    elif view.count('руб;') == 1:
                        stroka = view[view.find('руб;') - 20:view.find('руб;') + 4].split()
                        for j in range(len(stroka)):
                            stroka[j] = stroka[j].replace(',', '.')
                            stroka[j] = stroka[j].replace(' ', '')
                        if 'тыс.' in stroka or 'тыс.руб;' in stroka or 'тыс' in stroka:
                            f_k = 1
                        if 'млн.' in stroka or 'млн' in stroka:
                            f_kk = 1
                        stroka_num = ''
                        for j in range(len(stroka)):
                            try:
                                float(stroka[j])
                                stroka_num += stroka[j]
                            except:
                                continue
                        money += float(stroka_num)
                        if f_k:
                            kof = 1000
                        elif f_kk:
                            kof = 1000000
                        money *= kof

                    elif view.count('руб.') == 2:
                        border = view.find('руб.')
                        for _ in range(view.count('руб.')):
                            money1 = 0
                            stroka = view[border - 20:border + 4].split()
                            border = view[border + 4:].find('руб.') + border
                            for j in range(len(stroka)):
                                stroka[j] = stroka[j].replace(',', '.')
                                stroka[j] = stroka[j].replace(' ', '')
                            if 'тыс.' in stroka or 'тыс.руб.' in stroka or 'тыс' in stroka:
                                f_k = 1
                            if 'млн.' in stroka or 'млн' in stroka:
                                f_kk = 1
                            stroka_num = ''
                            for j in range(len(stroka)):
                                try:
                                    float(stroka[j])
                                    stroka_num += stroka[j]
                                except:
                                    continue

                            money1 += float(stroka_num)
                            if f_k:
                                kof = 1000
                            elif f_kk:
                                kof = 1000000
                            money1 *= kof
                            money += money1

                    list_money.append(money)

                    # Организационные причины
                    content = content.lower()
                    start_reasons1 = content.find('организационные причины')
                    start_reasons2 = content.find('прочие причины')
                    start_reasons3 = content.find('организационные причины аварии')

                    b = content.find('извлеченные уроки')
                    c = content.find('мероприятия по локализации')

                    reasons = []

                    if start_reasons1 != -1:
                        r_s = content[start_reasons1 + len('организационные причины'):]
                        double_ent = r_s[20:].find('\n\n')
                        r_s1 = r_s[:double_ent]
                        r_s1 = r_s1.replace('\n', '').replace(':', '')
                        if 'отсутствуют' in r_s1:
                            reasons.append('отсутствуют')
                        elif '2.1.' in r_s1:
                            r_s1 = r_s1.replace('2.1.', '')
                            if '2.2' in r_s1:
                                r_s1 = r_s1.split('2.2.')
                                reasons.append(r_s1[0])
                                if '2.3.' in r_s1[-1]:
                                    r_s1[-1] = r_s1[-1].split('2.3.')
                                    reasons.append(r_s1[-1][0])
                                    if '2.4.' in r_s1[-1][-1]:
                                        r_s1[-1][-1] = r_s1[-1][-1].split('2.4.')
                                        reasons.append(r_s1[-1][-1][0])
                                        if '2.5.' in r_s1[-1][-1][-1]:
                                            r_s1[-1][-1][-1] = r_s1[-1][-1][-1].split('2.5.')
                                            reasons.append(r_s1[-1][-1][-1][0])

                                        else:
                                            reasons.append(r_s1[-1][-1][-1])
                                    else:
                                        reasons.append(r_s1[-1][-1])
                                else:
                                    reasons.append(r_s1[-1])
                            else:
                                reasons.append(r_s1)

                        elif start_reasons3 != -1:
                            if 'отсутствуют' in r_s1:
                                reasons.append('отсутствуют')
                            elif '1.' in r_s1:
                                r_s1 = r_s1.replace('1.', '')
                                if '2.' in r_s1:
                                    r_s1 = r_s1.split('2.')
                                    reasons.append(r_s1[0])
                                    if '3.' in r_s1[-1]:
                                        r_s1[-1] = r_s1[-1].split('3.')
                                        reasons.append(r_s1[-1][0])
                                        if '4.' in r_s1[-1][-1]:
                                            r_s1[-1][-1] = r_s1[-1][-1].split('4.')
                                            reasons.append(r_s1[-1][-1][0])
                                            if '5.' in r_s1[-1][-1][-1]:
                                                r_s1[-1][-1][-1] = r_s1[-1][-1][-1].split('5.')
                                                reasons.append(r_s1[-1][-1][-1][0])
                                                if '6.' in r_s1[-1][-1][-1][-1]:
                                                    r_s1[-1][-1][-1][-1] = r_s1[-1][-1][-1][-1].split('6.')
                                                    reasons.append(r_s1[-1][-1][-1][-1][0])
                                                else:
                                                    reasons.append(r_s1[-1][-1][-1][-1])

                                            else:
                                                reasons.append(r_s1[-1][-1][-1])
                                        else:
                                            reasons.append(r_s1[-1][-1])
                                    else:
                                        reasons.append(r_s1[-1])
                                else:
                                    reasons.append(r_s1)
                            else:
                                reasons.append(r_s[len('аварии') + 2:r_s.find('\n\n')])

                    elif start_reasons2 != -1:
                        r_s = content[start_reasons2 + len('прочие причины'):]
                        double_ent = r_s.find('\n\n')
                        r_s1 = r_s[:double_ent]
                        r_s1 = r_s1.replace('\n', '').replace(':', '')
                        if 'отсутствуют' in r_s1:
                            reasons.append('отсутствуют')
                        elif '2.1.' in r_s1:
                            r_s1 = r_s1.replace('2.1.', '')
                            if '2.2' in r_s1:
                                r_s1 = r_s1.split('2.2.')
                                reasons.append(r_s1[0])
                                if '2.3.' in r_s1[-1]:
                                    r_s1[-1] = r_s1[-1].split('2.3.')
                                    reasons.append(r_s1[-1][0])
                                    if '2.4.' in r_s1[-1][-1]:
                                        r_s1[-1][-1] = r_s1[-1][-1].split('2.4.')
                                        reasons.append(r_s1[-1][-1][0])
                                        if '2.5.' in r_s1[-1][-1][-1]:
                                            r_s1[-1][-1][-1] = r_s1[-1][-1][-1].split('2.5.')
                                            reasons.append(r_s1[-1][-1][-1][0])

                                        else:
                                            reasons.append(r_s1[-1][-1][-1])
                                    else:
                                        reasons.append(r_s1[-1][-1])
                                else:
                                    reasons.append(r_s1[-1])
                            else:
                                reasons.append(r_s1)

                    if len(reasons) == 0:
                        reasons.append('отсутствуют')
                    list_org_reason.append(reasons)

            df = pd.DataFrame([[i, j, k, m1, n, o, p] for i, j, k, m, n, o, p in
                               zip(list_date_collapse, list_name_organisation, list_reason_collapse, list_org_reason,
                                   list_deaths,
                                   list_travm, list_money) for m1 in m],
                              columns=['Дата', 'Название', 'Причина', 'Организационные причины', 'Люди', 'Травмы',
                                       'Деньги'])
            df = df.sort_values(['Люди', 'Травмы', 'Деньги'], ascending=[False, False, False])
            df.to_excel(f"{self.path_report}\\Отчет.xlsx")

            self.popup_message.content.text = 'Создание отчета завершено'
            self.popup_message.open()

        def converter(self):
            wdFormatPDF = 17

            if not self.path_doc_files:
                self.popup_warning.content.text = 'Выберите директорию doc файлов'
                self.popup_warning.open()
                return

            if not self.path_pdf_files:
                self.popup_warning.content.text = 'Выберите директорию pdf файлов'
                self.popup_warning.open()
                return

            if not os.path.exists(f'{self.path_pdf_files}'):
                os.mkdir(f'{self.path_pdf_files}')

            for year in os.listdir(self.path_doc_files):

                if not os.path.exists(f'{self.path_pdf_files}\\{year}'):
                    os.mkdir(f'{self.path_pdf_files}\\{year}')

                word = comtypes.client.CreateObject('Word.Application')

                for file in os.listdir(f'{self.path_doc_files}\\{year}'):
                    path_pdf_full = f'{self.path_pdf_files}\\{year}\\{file[:file.rfind(".")]}.pdf'
                    path_doc_full = f'{self.path_doc_files}\\{year}\\{file}'
                    if not os.path.exists(path_pdf_full):
                        doc = word.Documents.Open(path_doc_full)
                        doc.SaveAs(path_pdf_full, FileFormat=wdFormatPDF)
                        doc.Close()

                word.Quit()

            self.popup_message.content.text = 'Конвертация завершена'
            self.popup_message.open()


    MainApp().run()

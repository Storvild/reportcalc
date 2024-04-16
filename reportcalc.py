import sys, os
import zipfile
import shutil
import io

import bs4
import lxml
import lxml.etree
import lxml.objectify
from bs4 import BeautifulSoup
from typing import Optional

print('Python:', sys.version)
curdir = os.path.dirname(__file__)
os.chdir(curdir)

TEMP_DIR = os.path.join(curdir,'ods_temp')
TEMP_CONTENT_DIR = os.path.join(TEMP_DIR,'content')
TEMP_CONTENT_FILEPATH = os.path.join(TEMP_CONTENT_DIR, 'content.xml')
#TEMPLATE_FILEPATH = os.path.join(curdir,'5.17 Справка о наличии ценностей, учитываемых на забалансовых счетах.ods')
TEMPLATE_FILEPATH = os.path.join(curdir,'TestReport001.ods')
RESULT_FILEPATH = os.path.join(TEMP_DIR, 'Result.ods')


def get_content(source_ods):
    """ Получение контента ods-файла в виде строки """
    with zipfile.ZipFile(source_ods, 'r') as odsfile:
        res = odsfile.read('content.xml')
        return res 

def write_to_file(content, source_ods, destination_ods):
    """ Запись контента в результирующий ods-файл """
    memfile_list = []
    # Считываем zip файл в память (массив memfile_list)
    with zipfile.ZipFile(source_ods, 'r') as odsfile:
        for item in odsfile.filelist:
            item = {'filename': item.filename, 'content': odsfile.read(item.filename), 'is_dir': item.is_dir()}
            memfile_list.append(item)
    # Создать папку для исходящего файла если ее не существует
    if not os.path.exists(os.path.dirname(destination_ods)):
        os.makedirs(os.path.dirname(destination_ods))
    # Записываем в новый zip-архив с новым файлом content.xml
    with zipfile.ZipFile(destination_ods, 'w', compression=zipfile.ZIP_DEFLATED) as odsfile: #compression=ZIP_STORED|ZIP_DEFLATED|ZIP_BZIP2|ZIP_LZMA ,compresslevel=(от 0 до 9 для ZIP_DEFLATED и ZIP_LZMA)
        for item in memfile_list:
            if item['filename']!='content.xml':
                odsfile.writestr(item['filename'], item['content'])
        odsfile.writestr('content.xml', content)
    return True
    
def run_file(filepath):
    """ Открытие файла связанной программой """
    #import subprocess
    code = os.startfile(filepath)
    return code




def change_content_etree(content):
    root = lxml.etree.fromstring(content)
    #for i in root.children('body'):
    #    print(i)
    #print(content)
    #print(dir(root))
    #print(dir(lxml.etree))


    #table_list = root.xpath('office:document-content/office:body/office:spreadsheet/table:table')
    # table_list = root.xpath('office:document-content', namespaces={'office': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'}) #http://base.google.com/ns/1.0
    # print(table_list)
    # for i in root.getchildren():
        # print(i)

    res = lxml.etree.tostring(root)  #, pretty_print=True
    print(res)
    return
    

def change_content_objectify(content):
    root = lxml.objectify.fromstring(content)
    for i in root.body.spreadsheet.getchildren():
        print(i, i.text, i.clear)
        print(dir(i))
    #print(dir(root.body.spreadsheet))
    return

def change_content_xpath(content):
    from copy import deepcopy
    root = lxml.etree.fromstring(content)
    #<office:document-content xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2">
    #print(root.nsmap)
    #table_list = root.xpath('office:document-content/office:body/office:spreadsheet/table:table', namespaces={'office':''})
    #table_list = root.xpath('office:document-content/office:body/office:spreadsheet', namespaces={'office':'urn:oasis:names:tc:opendocument:xmlns:table:1.0'})
    #print(root)
    #table_list = root.xpath('.//office:document-content/office:body/office:spreadsheet', namespaces={'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0','table':'urn:oasis:names:tc:opendocument:xmlns:table:1.0'})
    #table_list = root.xpath('office:body', namespaces={'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0','table':'urn:oasis:names:tc:opendocument:xmlns:table:1.0'})
    #table_list = root.findall('office:body/office:spreadsheet', namespaces={'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0','table':'urn:oasis:names:tc:opendocument:xmlns:table:1.0'})
    spreadsheet = root.xpath('office:body/office:spreadsheet', namespaces=root.nsmap)[0]
    #print(spreadsheet)
    print()
    table_list = root.findall('office:body/office:spreadsheet/table:table', namespaces=root.nsmap)
    for table in table_list:
        sheetname = table.attrib['{{{table}}}name'.format(**root.nsmap)]
        print(sheetname)
        if sheetname=='List1':   
            #spreadsheet.append(table.)
            new_table = deepcopy(table)
            new_table.attrib['{{{table}}}name'.format(**root.nsmap)] = 'List20'
            table.addnext(new_table)

            new_table2 = deepcopy(table)
            new_table2.attrib['{{{table}}}name'.format(**root.nsmap)] = 'List22'
            spreadsheet.append(new_table2)
            # Проверить может встречаться <text:p>текст<text:s>с тегом обозначающим пробел</text:p>
        #print(i.items())
        #print(dir(table))

    
    print()
    
    table1 = spreadsheet.xpath('table:table[@table:name="List1"]', namespaces=root.nsmap)[0]
    print(table1)
    #print(lxml.etree.iterparse(root))
    #print(dir(lxml.etree))
    print()
    textp = spreadsheet.xpath('.//text:p' , namespaces=root.nsmap)
    for i in textp:
        print(i.text, i.tag)
        
    print()
    exit()
    # Рекурсивный проход по дереву
    
    #table1 = spreadsheet.xpath('table:table[@table:name="табл_5_12"]', namespaces=root.nsmap)[0]
    for _, val in lxml.etree.iterwalk(table1):
        #print(k)
        print(val.text, val.tag)
        #print(type(val))
        #try:
        if val.text and '#field2#' in val.text:
            val.text = val.text.replace('#field2#', '@@@@@0005@@@@@')
            #print(val.text)
            print(val.text)
    #print(lxml.etree.tostring(table1))    

    #parent = table1.xpath('..')
    #print(parent)
    
    #table_list = root.findall('office:body/office:spreadsheet/table:table', namespaces=root.nsmap)
    #for table in table_list:
    #    print(table)
        
    #print(table_list)
    #table_list = root.xpath('office:document-content', namespaces={'office': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'}) #http://base.google.com/ns/1.0
    #for i in table_list.getchildren():
    #    print(i)
    #print(table_list)
    #print(root)
    return root

def mytemp():
    root = lxml.etree.fromstring('<root><table-cell value-type="string"><p>This text<s/>with Space<t/> as tag s and t</p></table-cell></root>')
    a = root.xpath('table-cell/p')[0].text
    print(a)
    
    for _, val in lxml.etree.iterwalk(root):
        print(val.text, val, val.tail)
        #print(val)
    pass


def old():
    # Создаем временную папку
    if not os.path.exists(TEMP_CONTENT_DIR):
        os.makedirs(TEMP_CONTENT_DIR)
        
    # Распаковываем контент
    with zipfile.ZipFile(TEMPLATE_FILEPATH, 'r') as odsfile:
        #a = odsfile.read('content.xml') # Чтение файла в память
        #print(a)
        #odsfile.extract('content.xml',path=TEMP_CONTENT_DIR)
        odsfile.extractall(path=TEMP_CONTENT_DIR)

    # Меняем content.xml    
    with open(TEMP_CONTENT_FILEPATH,'w') as fw:
        fw.write('123')

    #with zipfile.ZipFile(RESULT_FILEPATH, 'a') as odsfile:
    #    odsfile.write(TEMP_CONTENT_FILEPATH, arcname='content2.xml')

    # Копируем шаблон и запаковываем content.xml в него
    #shutil.copy2(TEMPLATE_FILEPATH, RESULT_FILEPATH)
    #with zipfile.ZipFile(RESULT_FILEPATH, 'a') as odsfile:
    #    odsfile.write(TEMP_CONTENT_FILEPATH, arcname='content2.xml')
        

    # Вариант запаковки через shutil.make_archive
    shutil.make_archive(RESULT_FILEPATH,'zip',TEMP_CONTENT_DIR)
    shutil.move(RESULT_FILEPATH+'.zip',RESULT_FILEPATH) # Убираем расширение .zip

    shutil.rmtree(TEMP_CONTENT_DIR) # Удаление временной директории со всем содержимым


class ReportCalc:
    def __init__(self):
        self.filepath = ''
        self.sheet_name = ''
        self.ods_xml_content: Optional[BeautifulSoup] = None

    def _get_content(self):
        """ Получение контента ods-файла в виде строки """
        with zipfile.ZipFile(self.filepath, 'r') as odsfile:
            res = odsfile.read('content.xml')
            return res

    def _get_sheetnames(self):
        """ Получение списка Листов """
        sheets = self.ods_xml_content.find('office:body').find('office:spreadsheet').find_all('table:table')
        res = [x['table:name'] for x in sheets]
        return res

    def _get_sheet(self, insheetname: str = '') -> Optional[bs4.element.Tag]:
        """ Получить один лист по имени
        :param insheetname: Имя листа. Если передан '', то возвращается первый лист
        :return: Тег листа <table:table>
        """
        sheets = self.ods_xml_content.find('office:body').find('office:spreadsheet').find_all('table:table')
        for s in sheets:
            if s['table:name'] == insheetname or insheetname == '':
                return s
        return None

    def find_row_id(self, insheet: bs4.element.Tag, intext: str) -> int:
        rows = insheet.find_all('table:table-row')
        for row_id, row in enumerate(rows):
            for item in row.descendants:
                if item.string and item.string == intext:
                    return row_id

    def find_row(self, insheet: bs4.element.Tag, intag_text: str, delete_tag: bool = True) -> bs4.element.Tag:
        rows = insheet.find_all('table:table-row')
        for row_id, row in enumerate(rows):
            for item in row.descendants:
                if item.string and item.string == intag_text:
                    if delete_tag:
                        item.decompose()
                    return row

    def replace_data(self, intag, fromstring, tostring):
        """ Заменяет один текст другим внутри тега intag рекурсивно """
        for tag in intag.descendants:
            if type(tag) != bs4.NavigableString and tag.string:
                tag.string = tag.string.replace(fromstring, tostring)
        return intag

    def insert_from_data(self, indata: list):
        """ Вставка данных """
        import copy
        sheet = self._get_sheet(self.sheet_name)
        #print(sheet.prettify())
        row = self.find_row(sheet, '#begintable', delete_tag=True)
        #row.decompose()  # Удаление тега
        #tag = row.copy()  # Копия
        #tag = row.clone()  # Копия
        #tag = row.extract()
        #sheet.append(tag)
        #sheet.insert_before(tag)
        #sheet.insert(0, tag)
        for x in indata:

            pass
        tag = copy.copy(row)
        self.replace_data(tag, 'строка', 'СТРОКА')
        sheet.append(tag)

        tag = copy.copy(row)
        self.replace_data(tag, 'Текст', 'ТЕКСТ')
        sheet.append(tag)

        row.decompose()

        print(sheet.prettify())

        #print(rows)

    def _test_replace_text(self, soup, old_text, new_text):
        """ Временно """
        for tag in soup.descendants:
            if isinstance(tag, bs4.NavigableString):
                if old_text in tag:
                    tag.replace_with(tag.replace(old_text, new_text))
        return soup


    def _test_change_data(self, sheetname='List2'):
        sheet1: BeautifulSoup = self._get_sheet(sheetname)
        tables = sheet1.find_all('table:table-row')
        for t in tables:
            cell_list = t.find_all('table-cell')
            if cell_list:
                for cell in cell_list:
                    # self._test_replace_text(cell, 'оформ', 'ОФОРМ')
                    # Рекурсивный проход черзе descendants
                    # for tag in cell.descendants:
                    #     if type(tag) != bs4.NavigableString and tag.string:
                    #         tag.string = tag.string.replace('оформ', 'ОФОРМ')
                    #     print(tag)
                    # Рекурсивный проход через find_all
                    for tag in cell.find_all(['text:p', 'text:span'], recursive=True):
                        if tag.string:
                            tag.string = tag.string.replace('строка', 'СТРОКА')
                    #    print(tag, '=>', tag.string)
                    # XML tag="<p> текст <i> с внутренним </i> тегом </p>" tag.string дает None
                    #   а без внутреннего тега <i> дает уже текст "текст с внутренним тегом"
                    pass

                        #if type(tag) == bs4.NavigableString:
                        #    tag.replace('строка', 'СТРОКА')
                        #    print(tag)
                        #if tag is not None and tag.string:
                        #    tag.string.replace_with("СТРОКА")
                        #    print(tag.string)
                # p_list = cell_list.find_all('text:p')
                # if len(p_list) > 0:
                #     for p in p_list:
                #         print(p.string)
                #         p.string = 'test'
                #         print('>>', p.string)

            #cell.text = 'test'
        #print(tables)
        print('\nXML Результат:')
        print(sheet1.prettify())

    def open_pattern(self, infilepath: str, insheetname: str='') -> bool:
        """ Обязательный метод, которым открывается шаблон"""
        self.filepath = infilepath
        self.ods_xml_content = BeautifulSoup(self._get_content(), 'xml')
        return True

    def _test_difficult_tag(self):
        inxml = '<main><p>Текст<i>с внутренним</i>тегом</p></main>'
        soup = BeautifulSoup(inxml, 'xml')
        p = soup.find_all('p')[0]
        p.string = 'TEST'
        print('p.string =', p.string)  # None


        print(soup)
        pass

if __name__ == '__main__':
    r = ReportCalc()
    #r.open_pattern(TEMPLATE_FILEPATH)
    r.open_pattern(r"e:\Temp\Без имени 1.ods")
    #print(r._get_sheet('List2'))
    #print(r._get_sheet())
    print(r.insert_from_data([{"fullname": "Наименование1", "quantity": 5}, {"fullname": "Наименование2", "quantity": 10.22}]))
    #print(r._test_change_data('Лист1'))
    #print(r._test_difficult_tag())
    #print(r._get_sheet('List2'))


    #content = get_content(TEMPLATE_FILEPATH) # Получаем контент
    #print(content)
    #change_content_etree(content)
    #change_content_objectify(content)
    #root = change_content_xpath(content)
    #content = lxml.etree.tostring(root)

    #mytemp()

    # Обрабатываем content

    #write_to_file(content, TEMPLATE_FILEPATH, RESULT_FILEPATH) # Записываем результирующий файл
    #run_file(RESULT_FILEPATH) # Запуск файла ods в OpenOffice


    print('Ok')



    
# WISH
# * Распаковка content.xml желательно в память
# * Распаковывать и запаковывать все желательно после обработки content.xml
#



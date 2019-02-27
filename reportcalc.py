import sys, os
import zipfile
import shutil
import io
import lxml
import lxml.etree
import lxml.objectify

print('Python:', sys.version)
curdir = os.path.dirname(__file__)
os.chdir(curdir)

TEMP_DIR = os.path.join(curdir,'ods_temp')
TEMP_CONTENT_DIR = os.path.join(TEMP_DIR,'content')
TEMP_CONTENT_FILEPATH = os.path.join(TEMP_CONTENT_DIR, 'content.xml')
TEMPLATE_FILEPATH = os.path.join(curdir,'5.17 Справка о наличии ценностей, учитываемых на забалансовых счетах.ods')
RESULT_FILEPATH = os.path.join(TEMP_DIR, 'Result.ods')


def get_content(source_ods):
    """ Получение контента в виде строки """
    with zipfile.ZipFile(source_ods, 'r') as odsfile:
        res = odsfile.read('content.xml')
        return res 

def write_to_file(content, source_ods, destination_ods):
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
    import subprocess
    code = os.startfile(filepath)


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

    res = lxml.etree.tostring(root) #, pretty_print=True
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
    
    # Рекурсивный проход по дереву
    
    #table1 = spreadsheet.xpath('table:table[@table:name="табл_5_12"]', namespaces=root.nsmap)[0]
    for _, val in lxml.etree.iterwalk(table1):
        #print(k)
        #print(val)
        #print(type(val))
        #try:
        if val.text and '#field2#' in val.text:
            val.text = val.text.replace('#field2#', '@@@@@0005@@@@@')
            #print(val.text)
            print(val.text)
    print(lxml.etree.tostring(table1))    

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
    
content = get_content(TEMPLATE_FILEPATH) # Получаем контент

#change_content_etree(content)    
#change_content_objectify(content)
#root = change_content_xpath(content) 
#content = lxml.etree.tostring(root)

mytemp()

# Обрабатываем content

#write_to_file(content, TEMPLATE_FILEPATH, RESULT_FILEPATH) # Записываем результирующий файл
#run_file(RESULT_FILEPATH) # Запуск файла ods в OpenOffice

        
print('Ok')
        
        
        
        
        
        
        
        
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

#print('ok')
    
# WISH
# * Распаковка content.xml желательно в память
# * Распаковывать и запаковывать все желательно после обработки content.xml
#



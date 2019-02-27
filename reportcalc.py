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
    
    
content = get_content(TEMPLATE_FILEPATH) # Получаем контент
root = lxml.etree.fromstring(content)
#root = lxml.objectify.fromstring(content)
#print(dir(root.body.spreadsheet))

table_list = root.xpath('office:document-content/office:body/office:spreadsheet/table:table', namespaces={'office': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0','table':'urn:oasis:names:tc:opendocument:xmlns:table:1.0'})
#table_list = root.xpath('office:document-content', namespaces={'office': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'}) #http://base.google.com/ns/1.0
print(table_list)


for i in root.getchildren():
    print(i)
	
#print(content)
#print(dir(root))
#print(dir(lxml.etree))



#res = lxml.etree.tostring(root) #, pretty_print=True
#print(res)
#content = res

# Обрабатываем content

#write_to_file(content, TEMPLATE_FILEPATH, RESULT_FILEPATH) # Записываем результирующий файл
#run_file(RESULT_FILEPATH) # Запуск файла ods в OpenOffice

        
print('OK')
        
        
        
        
        
        
        
        
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



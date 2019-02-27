import sys, os
import lxml
import zipfile
import shutil

curdir = os.path.dirname(__file__)
os.chdir(curdir)

TEMP_DIR = os.path.join(curdir,'ods_temp')
TEMP_CONTENT_DIR = os.path.join(TEMP_DIR,'content')
TEMP_CONTENT_FILEPATH = os.path.join(TEMP_CONTENT_DIR, 'content.xml')
TEMPLATE_FILEPATH = os.path.join(curdir,'5.17 Справка о наличии ценностей, учитываемых на забалансовых счетах.ods')
RESULT_FILEPATH = os.path.join(TEMP_DIR, 'Result.ods')

# Создаем временную папку
if not os.path.exists(TEMP_CONTENT_DIR):
    os.makedirs(TEMP_CONTENT_DIR)
    
zipfilepath = os.path.join(TEMPLATE_FILEPATH)
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

print('ok')
    
# WISH
# * Распаковка content.xml желательно в память
# * Распаковывать и запаковывать все желательно после обработки content.xml
#



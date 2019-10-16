import requests, time, os, platform, datetime, openpyxl, random, subprocess, tqdm, urllib
from bs4 import BeautifulSoup
#from collections import OrderedDict

url_sc = 'https://avto-yslyga.ru/wp-content/themes/auto/template-parts/check-inspection-handler.php'
url_ref = 'https://avto-yslyga.ru/proverit-tekhosmotr/'
KEY = "" 
internet = False
my_data = []

def creation_date(path_to_file):
    if platform.system() == 'Windows':
         t = os.path.getmtime(path_to_file)
         return datetime.datetime.fromtimestamp(t)  
    else:
        stat = os.stat(path_to_file)
        try:
            return stat.st_birthtime
        except AttributeError:
            return stat.st_mtime

def getKey(req):
    with open("key.txt", "wb") as code:
        code.write(req.content)
        code.close()

#while(len(NUMBER) < 8 or len(NUMBER) > 9):
    #NUMBER = str(input("Введите государственный номер ТС без пробелов: ")) 
def getGto(num):
    with requests.Session() as session:
        rk = session.get(url_ref) # Получаем страницу с формой логина
        if os.path.exists('key.txt'):
            date_now = datetime.date.today().strftime("%d.%m.%Y") 
            date_file = (creation_date('key.txt')).strftime("%d.%m.%Y")
            if(date_file != date_now):
                getKey(rk)
        else:
            getKey(rk)
        
        with open("key.txt", encoding='utf-8') as source:
            html = source.read()
            source.close()
        soup = BeautifulSoup(html, 'html.parser')    
        objs = soup.findAll(lambda tag: tag.name == 'p')
        KEY = objs[5].contents[1].get('value')
        dann = dict(regNumber = num, key = KEY) # Данные в виде словаря, которые будут отправляться в POST
        r = session.post(url_sc, dann, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2224.3 Safari/537.36'}) # Отправляем данные в POST, в session записываются наши куки  
    #print(r.text) # Дальше делайте с вашими данными все что захотите
        with open("result.txt", "wb") as text:
            text.write(r.content)
            text.close()

    with open("result.txt", encoding='utf-8') as data_text:
        data_html = data_text.read()
        data_text.close()

    sp = BeautifulSoup(data_html, 'html.parser')
    my_data = sp.findAll(lambda tag: tag.name == 'b')
    return my_data
    #print("Номер диагностической карты: ", my_data[0].text)
    #print("Марка ТС: ", my_data[1].text)
    #print("VIN ТС: ", my_data[2].text)
    #print("Номер кузова ТС: ", my_data[3].text)
    #print("Государственный номер: ", my_data[4].text)
    #print("Дата прохождения техосмотра: ", my_data[5].text)
    #print("Срок действия диагностической карты: ", my_data[6].text)
    #print("Оператор техосмотра: ", my_data[7].text)

while not internet:
    try:
        #subprocess.check_call(["ping", "-c 1", "www.google.com"])
        urllib.request.urlopen("http://google.com")
        print("Идет подключение к интернету")
        for i in tqdm.trange(100):
            time.sleep(0.01)
        internet = True
    except IOError:
        print("Отсутствует подключение к интернету!")

if internet == True and os.path.exists('spisok.xlsx'):
    print("Начата загрузка данных...")
    wb = openpyxl.load_workbook('spisok.xlsx')
    sht = wb['Лист1']
    sht['A1'] = "Гос. номер"
    sht['B1'] = "Номер диагностической карты"
    sht['C1'] = "Марка ТС"
    sht['D1'] = "VIN ТС"
    sht['E1'] = "Control"
    sht['F1'] = "Дата прохождения техосмотра"
    sht['G1'] = "Срок действия диагностической карты"
    sht['H1'] = "Оператор техосмотра"
    for i in range(2, sht.max_row+1):
        #print(sht.cell(row=i, column=1).value)
        gos_num = sht.cell(row=i, column=1).value
        get_data = getGto(gos_num)
        if len(get_data) > 0:
            sht.cell(row=i, column=2).value = get_data[0].text
            sht.cell(row=i, column=3).value = get_data[1].text
            sht.cell(row=i, column=4).value = get_data[2].text
            sht.cell(row=i, column=5).value = get_data[4].text
            sht.cell(row=i, column=6).value = get_data[5].text
            sht.cell(row=i, column=7).value = get_data[6].text
            sht.cell(row=i, column=8).value = get_data[7].text
            time.sleep(random.randint(150, 180)/10)
        #print(get_data)
        #wb.save('spisok.xlsx')
    wb.save('spisok.xlsx')
            
else:
    print("Нет списка техники для проверки!")

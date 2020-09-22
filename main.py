import project_chek_docs.dll as dll
import pandas as pd
import datetime
import  json


# Названия файлов
d = input('Введите дату контроля дд.мм.гггг')
DATE_in_File = pd.to_datetime(d,dayfirst=True)
PATH_FILE_Customers_bils_file = 'ОперативныйУчет 2017.xlsm'
MANAGER = 'Андрей'
PATH_FILE_manager_deals_file = 'СделкиМенеджеры.xlsx'
PATH_FILE_DOCS = 'Накладные.xlsx'

months_from_check={}
months = {'Январь':1, 'Февраль':2, 'Март':3, 'Апрель':4, 'Май':5, 'Июнь':6, 'Июль':7, 'Август':8, 'Сентябрь':9, 'Октябрь':10, 'Ноябрь':11, 'Декабрь':12}
now_month = datetime.datetime.now().month
now_year = datetime.datetime.now().year
errors={}
errors['list_n_find']=[]
errors['bill_n_find']=[]
errors['finans_err']=[]
errors['docs']=[]
errors['time_pay']=[]

# Определим классы
cbf = dll.Customers_bils_file(PATH_FILE_Customers_bils_file)
md = dll.manager_deals_file(PATH_FILE_manager_deals_file,MANAGER)
doc = dll.documents_control_file(PATH_FILE_DOCS)
work_file = dll.write_to_exel(PATH_FILE_manager_deals_file,MANAGER)

# Наченм работать

errors['duplicates'] = md.check_dupl()
print('Ошибки дублирования -' ,errors['duplicates'])
chek_list = md.list_for_check()


def get_rate(time_pay):
    if time_pay<=14:
        return 0.3
    else:
        return 0.25

for i,j in chek_list.iterrows():
    m=j['месяц'].title()
    y=str(now_year) if months[m]<=now_month else str(now_year-1)
    if not((m+y) in months_from_check):
        try:
            months_from_check[m+y]=cbf.get_sheet_month(m ,y)
        except:
            print(f'в строке {i} не найден месяц {m} в опреучете')
            errors['list_n_found'].append(f'месяц {m} в строке {i}')
            continue
    data_for_chek =  months_from_check[m+y]
    try:
        target_line = data_for_chek[data_for_chek['Номенклатурах']==j['Счет/доставка']].iloc[0]
    except:
        print(f"в строке {i} не найден счет {j['Счет/доставка']} в опреучете")
        errors['bill_n_find'].append(f"в строке {i} не найден счет {j['Счет/доставка']} в опреучете")
        continue
    if  abs(target_line['Выручка']-  j['Выручка'])>1 or abs(target_line['Маржа'] - j['Маржа'])>1 or abs(target_line['Ст. Закупки']- j['Ст. Закупки'])>1 or abs(target_line['Поставщик/Откат']- j['Откат'])>1:
        print(f"В строке {i}, по счету {j['Счет/доставка']} не соответсвут финасовые показатели")
        errors['finans_err'].append(f"В строке {i}, по счету {j['Счет/доставка']} не соответсвут финасовые показатели")
        continue
    if  not doc.check_docs(j['Счет/доставка']):
        print(f"В строке {i}, по счету {j['Счет/доставка']} не найдены подписанные документы")
        errors['docs'].append(f"В строке {i}, по счету {j['Счет/доставка']} не найдены подписанные документы")
        continue
    try:
        time_pay = -(pd.to_datetime(j['накладная'][-10:],dayfirst=True)  - j['Дата оплаты']).days
    except:
        print(f"В строке {i}, по счету {j['Счет/доставка']} невозможно определить дату платежа")
        errors['time_pay'].append(f"В строке {i}, по счету {j['Счет/доставка']} невозможно определить дату платежа")
        continue
    work_file.rec(i,get_rate(time_pay),DATE_in_File)
    print(f"Строка {i} по счету {j['Счет/доставка']} проверен и занесен")
print(errors)
with open('errors.txt','w',encoding='utf-8') as f:
    f.write(json.dumps(errors,ensure_ascii=False))
work_file.clos()
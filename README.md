Механизм проверки для начисления мотивации менеджеру по продажам
Для расчета своей заработной платы, менеджер по продажам ведет файл в Эксель. В котором указывает каждую сделку и е параметры (номер счета, накладные , выручка, стоимость закупки, маржа, дата оплаты и тд) Менеджер может получить свое вознаграждение за сделки, лишь после того, как сделка будет оплачена и он предоставить оригиналы подписанных клиентом финансовых документов. Назовём этот файл "Расчеты менеджера" Так же есть общий файл фирмы в Экселе, где указаны все сделки, который ведет проверяет специальный человек. Назовем его "Оперативный учет" Ранее процесс был такой: Менеджер заключает сделки, указывает их в своем файле, собирает первичные документы, ждет оплату, после этого передает все накладные другому специальному человеку. Он проверяет каждую накладную, сверяет записи финансовых показателей сделки в файле менеджера с файлом опреучета. Если все хорошо, ставит пометку в виде даты проверки, а так же определяет процент к начислению менеджера в зависимости от финансовой дисциплины контрагента. Далее менеджер получает свою зарплату, документы уходят в бухгалтерию, где бухгалтер их заносит в 1с и указывает галочку, что оригиналы получены. Человек который проверяет файл менеджера и его документы тратит на каждого несколько дней :)
Для начала я им предложил все сделать в одной БД и интерфейс к ней, но отказались так как им так привычнее и удобнее. Сделал следующее:
1.	Менеджер заносит свою сделку в свой файл (как и было)
2.	Далее сразу отдает документы в бухгалтерию, где бухгалтер их проверяет и делает отметку о наличии оригинала (таким образом вторую проверку мы избегаем)
3.	Далее при наступлении дня расчета зарплаты, мы берем файл менеджера, файл оперативного учета, а так же копируем отчет из 1С в эксель (где указано получен первичный документ или нет), файл "Накладные". Ложем их в одну папочку, запускаем скрипт и система за 1 минуту все сама отрабатывает
Данные из фалов загружаем в pandas, там с ними работаем, далее открываем и записываем нужные данные средствами openpyxl. Так же формируем файл с ошибками, где подробно указано, если что то с чем то не совпадает


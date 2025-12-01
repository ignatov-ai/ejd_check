#===============================================================================
#
# Для работы скрипта необходимы библиотеки:
#
# pip install lxml
# pip install pycurl-7.43.0.5-cp39-cp39-win_amd64.whl
# python -m pip install -U pip setuptools
# python -m pip install grab
# pip install pyquery
# pip install browser-cookie3
#
# !!! Библиотеку pycurl лучше ставить именно из указанного whl-контейнера,
#     который предварительно следует скачать на свой комп.
#     Например, с сайта https://www.lfd.uci.edu/~gohlke/pythonlibs/#pycurl
#     Фрагмент "cp39-cp39" в названии файла означает версию установленного
#     на компьютере Python`а. Если установлена 64-битная версия Python`а,
#     то следует ставить "win_amd64"-версию библиотеки.
#     В данный контейнер уже включен файл curl.exe, необходимый для корректной
#     работы библиотеки.
#
# !!! Скрипт тестировался в версии Python 3.9.x - 3.11.x
#     Вышеназванная библиотека pycurl указанной версии может некорректно
#     работать в других версиях Python - в этом случае установите библиотеку
#     pycurl совместимой версии, скачаной с вышеназванного сайта:
#     https://www.lfd.uci.edu/~gohlke/pythonlibs/#pycurl
#
# !!! В whl-контейнере для 11-ой версии Python`а нет самого файла curl.exe -
#     его можно взять из предыдущих версий.
#
#===============================================================================
#
# Консольное приложение!
# Т.е., данный скрипт запускается из командной строки.
#
# Например, из командной строки любого файлового менеджера или можно открыть
# окно командного процессора cmd:
# - в Windows из меню выбрать "Запустить", ввести cmd - откроется окно.
# - с помощью команд вида ">cd Каталог" перейти в каталог со скриптом
# - запустить скрипт (примеры см. ниже)
#
#===============================================================================
#
# ВНИМАНИЕ!!!
# В данном скрипте авторизация в ЭЖД производится на данных из куков браузера.
# Стабильно работает этот метод только в браузерах FireFox.
# Перед запуском скрипта необходимо в одном из вышеназванных браузеров
# войти в ЭЖД, перейти в старый интерфейс админки "Организация обучения",
# открыть справочник "Кадры" в меню "Общее образование".
# После этого, не выходя из ЭЖД и не закрывая браузер, можно запустить скрипт.
# Наиболее стабильно данный метод работает с браузером FireFox.
# Иногда при использовании браузера Opera при запуске скрипта может быть
# выдано сообщение о необходимости выполнения вышеназванных шагов - просто
# запустите скрипт еще раз через 20-30 секунд.
#
#===============================================================================

import sys, os, time

import browser_cookie3
import grab
from grab import Grab

classes_korp_8 = [  "5-А-З", "5-Б-З", "5-В-З", "5-Д-З", "5-Е-З", "5-З-З", "5-И-З", "5-К-З", "5-Л-З", "5-М-З", "5-Ю", "5-Я",
                    "6-А-З", "6-Б-З", "6-В-З", "6-Д-З", "6-Е-З", "6-З-З", "6-И-З", "6-Л-З", "6-Ф", "6-Ю", "6-Я",
                    "7-А-З", "7-Б-З", "7-В-З", "7-Д-З", "7-Ц", "7-Ч", "7-Ш", "7-Э", "7-Ю", "7-Я",
                    "8-А-З", "8-Б-З", "8-В-З", "8-Д-З", "8-Ц", "8-Ч", "8-Ш", "8-Э", "8-Ю", "8-Я",
                    "9-А-З", "9-Б-З", "9-В-З", "9-Д-З", "9-Ц", "9-Ч", "9-Ш", "9-Э", "9-Ю", "9-Я",
                    "10-Ф", "10-Ц", "10-Ч", "10-Ш", "10-Э", "10-Я",
                    "11-Ф", "11-Ц", "11-Ч", "11-Ш", "11-Э", "11-Я"]

# ===============================================================================
#
# Сохранение расширенных журналов классов.
#
# В качестве обязательного параметра задается имя папки для сохранения
#   архива со скачанными журналами. В абсолютной или относительной форме.
# В качестве необязательного параметра можно задать номер параллели или
#   название класса. Например, 10 или 10-А.
#   В первом случае будут скачаны журналы классов только заданной параллели.
#   Во втором случае - журналы только заданного класса.
#
# Примеры:
#
# >roa_save_jrn_ext jrn 6-А
# В папке "jrn", которая находится в папке, откуда запускается скрипт, будут 
#   сохранены xlsx-файлы по каждому журналу 6-А класса.
#
# >roa_save_jrn_ext "C:\temp\jrn" 6
# В папке "C:\temp\jrn" будут сохранены xlsx-файлы с журналами всех 6-х классов
#   и метагрупп по 6-ой параллели.
#
# >roa_save_jrn_ext jrn
# В папке "jrn", которая находится в текущей папке, будет создан один zip-архив
#   с журналами всех классов школы, включая все метагруппы.
#
#===============================================================================
#===============================================================================

class dn_Auth:

      domain = ""; base = ""
      web = None
      # pid = "89152360386" # user profile id
      pid = "ignatov-ai@yandex.ru" # user profile id
      sid = "" # school id
      aid = "13" # academic year id
      curr_aid = ""
      dict_th = None; dict_subj = None; dict_bld = None; dict_room = None

      def __init__(self, dn="test", timeout=15, conn_tm=10, aid=""):
          self.timeout = timeout
          self.conn_tm = conn_tm

          if dn.lower() == "work":
             self.domain = "dnevnik.mos.ru"
             self.base   = "https://dnevnik.mos.ru/"
          else:
             self.domain = "dnevnik-test.mos.ru"
             self.base   = "https://dnevnik-test.mos.ru/"

          self.free()
          self.aid = aid

      def __del__(self):
          self.free()

      def free(self):
          if self.web != None:
             self.web.reset()
             self.web = None
          self.pid = ""
          self.dict_th   = {}
          self.dict_subj = {}
          self.dict_bld  = {}
          self.dict_room = {}

      def set_aid(self, aid=""):
          if not self.web: return("")

          aid = aid.strip()
          if (not aid) or (aid == "0"):
             aid = self.curr_aid
          elif aid[0] == "-":
             shift = aid[1:]
             aid = self.curr_aid
             if shift.isdigit() and int(shift):
                shift = int(shift)
                self.web.go(self.base + "core/api/academic_years")
                if (self.web.doc.code == 200) and \
                   (len(self.web.doc.json) > shift):
                   cy = False
                   for y in sorted(self.web.doc.json,             \
                                   key=lambda y: y["begin_date"], \
                                   reverse=True):
                       if y["current_year"]:
                          cy = True
                       elif cy:
                          shift -= 1
                          if not shift:
                             aid = str(y["id"])
                             break

          self.aid = aid
          self.web.config['cookies']['aid'] = self.aid
          self.web.cookies.set('aid', self.aid, \
                               "."+ self.base.split("://")[-1].split("/")[0])

          return(self.aid)

      def reset(self):
          # Данный метод рекомендуется вызывать перед любым запросом, если
          # перед этим был POST/PUT/DELETE запрос.
          if self.web:
             self.web = self.web.clone()
             self.web.setup(headers={'Content-Type':'application/json;charset=UTF-8'})

      def fetch(self, uri, page_size="per_page", page_num="page", pages=1000):
          # В параметре page_size можно указать размер страницы -
          # "per_page=1000". В этом случае данный параметр будет добавлен
          # в запрос без изменений, "как есть".
          # Если параметр pages=0, то будет сформирован обычный запрос
          # без пагинации.
          if not self.web or not uri or (pages < 0):
             return None

          self.reset()

          paginate = "?" if "?" not in uri else "&"
          if pages == 0:
             pages = 1
             paginate = ""
          elif page_size:
             if "=" in page_size:
                paginate += page_size
             else:
                paginate += page_size +"=1000"
             if page_num: paginate += "&"+ page_num +"="
          elif page_num:
             paginate += page_num +"="
          data = []
          for page in range(1, pages+1):
              if len(paginate):
                 self.web.go(uri + paginate + str(page))
              else:
                 self.web.go(uri)

              if self.web.doc.code != 200:
                 return None
              elif (not self.web.doc.json) or (not len(self.web.doc.json)):
                 break
              else:
                 data.extend(self.web.doc.json)
          return data

      def login(self):
          if self.web: return True

          # Подготовка к подключению
          self.web = Grab(connect_timeout = self.conn_tm, timeout = self.timeout)
          self.web.config['common_headers']['Accept-Language'] = 'ru-RU'
          self.web.config['common_headers']['Accept'] = 'application/json,'+ \
                         self.web.config['common_headers']['Accept']

          # Запрос страницы входа для получения общих куков
          self.web.go(self.base)
          if (self.web.doc.code != 200):
             print("Ошибка доступа к сайту \""+ self.base +"\"!!!")
             self.free()
             return False

          # Поиск куков авторизации в ЭЖД и их сохранение
          cookies = None
          try:
            cookies = browser_cookie3.firefox(domain_name = self.domain)
            if str(cookies).find("is_auth=true") == -1:
               cookies = None
          except browser_cookie3.BrowserCookieError:
            pass

          if not cookies:
             try:
               cookies = browser_cookie3.opera(domain_name = self.domain)
               if str(cookies).find("is_auth=true") == -1:
                  print("Откройте браузер FireFox или Opera, авторизируйтесь в ЭЖД,\nперейдите в справочник \"Кадры\" и запустите программу снова!")
                  self.free()
                  return False
             except browser_cookie3.BrowserCookieError:
               print("Используйте браузеры FireFox или Opera!")
               self.free()
               return False

          self.web.setup(headers={'Content-Type': 'application/json;charset=UTF-8'})
          self.web.setup(connect_timeout=self.conn_tm, timeout=self.timeout)
          self.web.config["cookies"] = {}
          for cookie in cookies:
              line = str(cookie)[8:]
              pos = line.find(" ") 
              if pos == -1: continue
              (cname, cvalue) = line[:pos].split("=", maxsplit=1)
              self.web.config["cookies"][cname] = cvalue

          if "profile_id" not in self.web.config["cookies"]:
             print("Ошибка авторизации!!!")
             self.free()
             return False

          self.pid = str(self.web.config["cookies"]["profile_id"])

          self.web.config['common_headers']['Auth-Token'] = \
                         self.web.config["cookies"]["auth_token"]
          self.web.config['common_headers']['Profile-Id'] = self.pid

          # Ид текущего учебного года
          start_aid = self.aid
          # self.aid = str(self.web.config["cookies"]["aid"])
          self.aid = "13"
          self.curr_aid = self.aid or "13"
          self.set_aid(start_aid if start_aid else self.curr_aid)

          # Определение ид школы
          self.web.go(self.base + "core/api/schools")
          if (self.web.doc.code != 200):
             print("Ошибка авторизации ("+ str(self.web.doc.code) +")!!!")
             self.free()
             return False

          self.sid = self.web.doc.json[0]["id"]
          
          return True

#===============================================================================
#===============================================================================

if len(sys.argv) < 2:
   print("Запуск: save_jrn_ext path_folder [par_or_class]")
   print("path_folder  - путь к папке для сохранения журналов")
   print("  Одна точка - сохранение в текущую папку.")
   print("par_or_class - номер параллели или имя класса.")
   print("  Может быть опущен - в этом случае будут скачаны журналы всех классов.")
   exit(-1)

path_folder = sys.argv[1].strip()
if path_folder == ".":
   path_folder = ""
elif not os.path.isdir(path_folder):
   print("Указанная папка не существует - задайте путь к существующей папке!")
   exit(-3)
elif path_folder[-1] != "\\":
   path_folder += "\\"

cl_level = 0
cl_name = ""
if len(sys.argv) > 2:
   cl_name = sys.argv[2].strip().upper()
   if len(cl_name) < 3:
      cl_level = int(cl_name)
      cl_name = ""

# Авторизация в ЭЖД
dn = dn_Auth("work", timeout=30)
if (dn is None) or not dn.login():
   print("Ошибка входа в ЭЖД!")
   exit(-4)

cDate = time.localtime()
if (cDate.tm_mon > 7):
   sDate = str(cDate.tm_year) +"-09-01T00:00:00.000Z"
   eDate = str(cDate.tm_year+1) +"-05-31T23:59:59.999Z"
else:
   sDate = str(cDate.tm_year-1) +"-09-01T00:00:00.000Z"
   eDate = str(cDate.tm_year) +"-05-31T23:59:59.999Z"

# Запрос списка всех классов
dn.web.go(dn.base +"core/api/class_units?with_home_based=true&academic_year_id="+ dn.aid)
if dn.web.doc.code != 200:
   print("Ошибка получения списка классов! ("+ str(dn.web.doc.code) +")")
   exit(-5)

grpfail = open(path_folder +"jrn_ext_err.log","w")
grp_fail = []

tmap = "".maketrans(":<>*/\\\"'","-[]#||``",".?")

# Обходим классы
jrn_count = 0
for cl in dn.web.doc.json:
    class_level = cl.get("class_level_id", 0)
    if not (0 < class_level < 12): continue

    class_name  = cl.get("name", "").upper()

    # скачивание только классов корпуса №8
    # if class_name not in classes_korp_8: continue

    if cl_level:
       if class_level != cl_level: continue
    elif cl_name:
       if class_name != cl_name: continue

    class_id = cl.get("id", 0)

    # Запрашиваем список групп класса
    print(class_name)
    grps = dn.fetch(dn.base +"jersey/api/groups?class_unit_id="+ str(class_id) + \
                    "&academic_year_id="+ dn.aid +"&with_lessons_only=true")
    if dn.web.doc.code != 200:
       print("Ошибка получения списка групп класса! ("+ str(dn.web.doc.code) +")")
       continue

    # Обходим группы
    for grp in grps:
        grp_id = grp.get("id", 0)
        grp_name  = grp.get("name", "")
        if (not grp_id) or (not grp_name): continue

        subj_name = str(grp.get("subject_name", "")).translate(tmap)

        grp_name = grp_name.replace("/","_")

        ok = False
        for k in [1,2]:
          try:
            if k == 2: time.sleep(1)
            dn.web.go(dn.base +"export/journal.xlsx?group_ids="+ str(grp_id) + \
                      "&extended=true&start_at="+ sDate +"&stop_at="+ eDate)
            if (dn.web.doc.code == 200) and (dn.web.doc.download_size > 2000):
               print(grp_name)

               # dn.web.doc.save(path_folder + ";" + class_name +";"+ subj_name +";"+ grp_name +".xlsx")
               dn.web.doc.save(path_folder + "\\" + str(class_level) + "\\" + class_name +";"+ subj_name +";"+ grp_name +".xlsx")

               jrn_count += 1
               ok = True
               break
          except grab.error.GrabTimeoutError:
            pass
        if not ok:
           grp_fail.append([grp_id, class_name, grp_name, subj_name])
           grpfail.write(str(class_level) + "\\" + class_name + ";" + subj_name + ";" + grp_name + ".xlsx")
           grpfail.flush()

# Повторяем сохранение "сбойных" журналов
if len(grp_fail):
   print("Повторное сохранение 'сбойных' журналов ("+ str(len(grp_fail)) +")")
   grpfail.write("\n=== Повтор ==========\n")

for i in range(len(grp_fail)-1,-1,-1):
    ok = False
    for k in [1,2]:
      try:
        if k == 2: time.sleep(1)
        dn.web.go(dn.base +"export/journal.xlsx?group_ids="+ str(grp_fail[i][0]) + \
                  "&extended=true&start_at="+ sDate +"&stop_at="+ eDate)
        if (dn.web.doc.code == 200) and (dn.web.doc.download_size > 2000):
           print(grp_fail[i][2])
           dn.web.doc.save(path_folder + "\\" + str(class_level) + "\\" + grp_fail[i][1] + " " + grp_fail[i][3] + "(" + grp_fail[i][2] + ").xlsx")
           jrn_count += 1
           del(grp_fail[i])
           ok = True
           break
      except grab.error.GrabTimeoutError:
        pass
    if not ok:
       grpfail.write(grp_fail[i][1] +" "+ grp_fail[i][3] +"("+ grp_fail[i][2] +").xlsx\n")
       grpfail.flush()

if len(grp_fail):
   print("Не удалось сохранить журналы групп:")
   grpfail.write("\n=== Итого ==========\n")
   for grp in grp_fail:
       print(grp[2])
       grpfail.write(grp[1] +" "+ grp[3] +"("+ grp[2] +").xlsx\n")

del(dn)
grpfail.close()
if not len(grp_fail): os.remove(grpfail.name)

if not jrn_count:
   if cl_name:
      print("Класс не найден!")
   else:
      print("Не скачано ни одного журнала!")
   exit(-100)
else:
   print("Скачано журналов: "+ str(jrn_count))
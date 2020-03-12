from selenium import webdriver
##import openpyxl
from openpyxl import *
import time

EGRN = None
status = 0 # состояние сайта
row = 0
wb = None
ws = None
  
##------------------------------------------------------------##
def init():
  global EGRN
  EGRN = webdriver.Chrome()
  ##  EGRN.implicitly_wait(1)
  EGRN.set_page_load_timeout(10)
  EGRN.get("https://rosreestr.ru/wps/portal/p/cc_present/ir_egrn")

##------------------------------------------------------------##
def init_pyxl():
  global wb
  global ws
  try:
    wb = load_workbook(filename = 'dn.xlsx')
  except:
    wb = Workbook()
  ws = wb.worksheets[0]
  if ws.cell(row = 1, column = 1).value == None:
    ws.cell(row = 1, column = 1).value = "Номер запроса"
    ws.cell(row = 1, column = 2).value = "Дата запроса"
    ws.cell(row = 1, column = 3).value = "Статус запроса"
    wb.save(filename = 'dn.xlsx')

##------------------------------------------------------------##
def dn(d):
  global row
  global wb
  global ws
  global EGRN
  table = EGRN.find_element_by_css_selector("table.v-table-table")
  for tr in table.find_elements_by_css_selector("tr"):
    NZ = tr.find_elements_by_css_selector("td")[0].get_attribute("innerText")
    DZ = tr.find_elements_by_css_selector("td")[1].get_attribute("innerText")
    SZ = tr.find_elements_by_css_selector("td")[2].get_attribute("innerText")
    LZ = tr.find_elements_by_css_selector("td")[3].find_elements_by_css_selector("a")

    found = False
    for r in range(2, ws.max_row + 1):
      if ws.cell(row = r, column = 1).value == NZ:
        found = True
        break

    if not found:
      row = ws.max_row + 1
    else:
      row = r
    status = ws.cell(row = row, column = 3).value
    if  status != "Завершена":
      ws.cell(row = row, column = 1).value = NZ
      ws.cell(row = row, column = 2).value = DZ
      ws.cell(row = row, column = 3).value = SZ
      if len(LZ) > 0 and d in DZ :
        LZ[0].click()
  wb.save(filename = 'dn.xlsx')
  lastdate = time.strptime(DZ[0:10], "%d.%m.%Y")
  d1 = time.strptime(d, "%d.%m.%Y")
  return  lastdate >= d1

def dnd(d, p1 = 1, p2 = 50):
  global EGRN
  init_pyxl()
  table = EGRN.find_element_by_css_selector("table.v-table-table")
  pm = EGRN.find_elements_by_css_selector("div.v-horizontallayout.v-horizontallayout-pageManagement.pageManagement > div > div")
  first_page = pm[1].find_element_by_xpath("div/div")
  is_first_page  = first_page.get_attribute("aria-pressed") != "false"
  if not is_first_page:
    first_page.click()
  while not is_first_page:
    time.sleep(1)
    is_first_page  = first_page.get_attribute("aria-pressed") != "false"
  print("Первая страница ОК")

  pm = EGRN.find_elements_by_css_selector("div.v-horizontallayout.v-horizontallayout-pageManagement.pageManagement > div > div")
  activepage = int(pm[3].find_element_by_css_selector("div.v-horizontallayout div.v-label.v-label-undef-w").get_attribute("innerText"))
  next_page = pm[4].find_element_by_css_selector("div")

  status = True
  if activepage in range(p1, p2 + 1):
    while dn(d) and status:
      print(f"Обработана странца {activepage}")
      old_activepage = activepage
      next_page.click()
      while old_activepage == activepage:
        time.sleep(2)
        try:
          pm = EGRN.find_elements_by_css_selector("div.v-horizontallayout.v-horizontallayout-pageManagement.pageManagement > div > div")
          activepage = int(pm[3].find_element_by_css_selector("div.v-horizontallayout div.v-label.v-label-undef-w").get_attribute("innerText"))
          next_page = pm[4].find_element_by_css_selector("div")
          status = True
        except:
          activepage = 9999
          status = False
  print("Всё готово")        
        



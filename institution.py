from openpyxl import workbook, load_workbook
import re
import bs4
import time
import random

wb_uasys = load_workbook(r"File Location")
wb_data_grab = load_workbook(r"File Location")
ws_uasys = wb_uasys["Name of Worksheet"]
ws_data_grab = wb_data_grab["Name of Worksheet"]
# Get INST_PO_BOX_LINE for PRIMARY_INSTITUTION_NAME from LocationName -> Address
for cell in ws_uasys['U']:
    PRIMARY_INSTITUTION_NAME = str(cell.value).upper()
    print("----------------------------------")
    print(PRIMARY_INSTITUTION_NAME)
    for grab in ws_data_grab['D']:
        Location_Name = str(grab.value)
        if Location_Name.upper() == PRIMARY_INSTITUTION_NAME:
            Address = str(ws_data_grab['H' + str(grab.row)].value)
            if Address.find('P.O'):
                found = re.search("x(.+?),", Address)
                if not found:
                    continue
                else:
                    number_POBOX = found.group(1)
                    INST_PO_BOX_LINE = str('PO Box' + str(number_POBOX))
                    print('Found: ' + INST_PO_BOX_LINE)
                    ws_uasys['X' + str(cell.row)].value = INST_PO_BOX_LINE
# If INST_COUNTRY_CODE is blank then assign USA
for cell in ws_uasys['AA']:
    try:
        if cell.value is None:
            ws_uasys['AA' + str(cell.row)].value = "USA"
    except AttributeError:
        print('cell is read only!')
# Get INST_ESTABLISHED_DATE for PRIMARY_INSTITUTION_NAME from Google search
print('Looking up Institution established dates.........')
for cell in ws_uasys['V']:
    PRIMARY_INSTITUTION_NAME = str(cell.value).upper()
    print(PRIMARY_INSTITUTION_NAME + ' was founded:')

    proxy_username = "USER_NAME"
    proxy_password = "PASSWORD"
    proxy_url = "http://proxy.scrapingbee.com"
    proxy_port = 8886
    options = {
        "proxy": {
            "http": f"https://{proxy_username}:{proxy_password}@{proxy_url}:{proxy_port}",
            "verify_ssl": False,
        },
    }

    url = 'https://google.com/search?q=' + '"' + str(PRIMARY_INSTITUTION_NAME) + '"' + ' / Founded'
    driver = webdriver.Chrome(
        executable_path=r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        seleniumwire_options=options,
    )
    driver.get(url)
    wait = random.randrange(1, 10)
    time.sleep(wait)
    request_result = driver.page_source
    driver.quit()
    web_data = bs4.BeautifulSoup(request_result, "html5lib")
    try:
        DATE = web_data.find('div', class_='Z0LcW t2b5Cf').text
        INST_ESTABLISHED_DATE = DATE
        print(INST_ESTABLISHED_DATE)
        ws_uasys['AG' + str(cell.row)].value = str(INST_ESTABLISHED_DATE) + '-01-01'
    except AttributeError:
        print("----------------------------------")
        print('NoneType for: ' + str(cell.value))
print('Done!')
wb_uasys.save(r"File Location")

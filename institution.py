from openpyxl import load_workbook
import undetected_chromedriver as uc
import re
import ssl
import bs4
import time
import random

# Change Add .xlsx
wb_uasys = load_workbook(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Missouri Educational Institutions 2023-05-26.xlsx")
wb_data_grab = load_workbook(r"C:\Users\Wayne Cole\Downloads\Work Stuff\AccreditationData.xlsx")
wb_nces_grab = load_workbook(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Data_3-14-2023---623.xlsx")
# Change
ws_uasys = wb_uasys["All Missouri Institutions"]
ws_data_grab = wb_data_grab["InstituteCampuses"]
ws_nces_grab = wb_nces_grab["Data_3-14-2023---623"]

# If PRIMARY_ORGANIZATION_ID is blank then assign the cell AutoGen
for cell in ws_uasys['Q']:
    try:
        if cell.value is None:
            ws_uasys['Q' + str(cell.row)].value = "AutoGen"
    except AttributeError:
        print(cell + ' is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')

# Get INST_PO_BOX_LINE for PRIMARY_INSTITUTION_NAME from LocationName -> Address
for cell in ws_uasys['U']:
    PRIMARY_INSTITUTION_NAME = str(cell.value).upper()
    cell_prev = int(cell.row) - 1
    try:
        if cell_prev != 0 and PRIMARY_INSTITUTION_NAME != ws_uasys['U' + str(cell_prev)].value.upper():
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
    except AttributeError:
        print("----------------------------------")
        print('NoneType for: ' + str(cell.value))
    except TypeError:
        print('NoneType')
    except:
        print('Unknown error')

# If INST_COUNTRY_CODE is blank then assign USA
for cell in ws_uasys['AA']:
    try:
        if cell.value is None:
            ws_uasys['AA' + str(cell.row)].value = "USA"
    except AttributeError:
        print('cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')

# Get INST_ESTABLISHED_DATE for PRIMARY_INSTITUTION_NAME from Google search
# work on denied access and headless evasion
print('Looking up Institution established dates.........')
for cell in ws_uasys['U']:
    PRIMARY_INSTITUTION_NAME = str(cell.value).upper()
    found_word1 = PRIMARY_INSTITUTION_NAME.find('BARBER')
    found_word2 = PRIMARY_INSTITUTION_NAME.find('BEAUTY')
    found_word3 = PRIMARY_INSTITUTION_NAME.find('HAIR')
    found_word4 = PRIMARY_INSTITUTION_NAME.find('SALON')
    found_word5 = PRIMARY_INSTITUTION_NAME.find('SPA')
    if found_word1 or found_word2 or found_word3 or found_word4 or found_word5 < 0:
        try:
            cell_prev = int(cell.row) - 1
            if cell_prev != 0 and PRIMARY_INSTITUTION_NAME != ws_uasys['U' + str(cell_prev)].value.upper():
                print(PRIMARY_INSTITUTION_NAME + ' was founded:')
                if ws_uasys['AF' + str(cell.row)].value is None:
                    ssl._create_default_https_context = ssl._create_unverified_context
                    chrome_options = uc.ChromeOptions()

                    url = 'https://google.com/search?q=' + '"' + str(PRIMARY_INSTITUTION_NAME) + '"' + ' / Founded'
                    driver = uc.Chrome(options=chrome_options)
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
                        ws_uasys['AF' + str(cell.row)].value = str(INST_ESTABLISHED_DATE) + '-01-01'
                        # change this save location between states
                        wb_uasys.save(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Illinois Educational Institutions 2023-05-26.xlsx")
                    except AttributeError:
                        print("----------------------------------")
                        print('NoneType for: ' + str(cell.value))
                    except TypeError:
                        print('Cell is read only!')
                    except:
                        print('Unknown error')
        except TypeError:
            print('That was a merged or empty cell skipping......')
        except AttributeError:
            print('NoneType for this cell')
        except:
            print('unknown error')

# Check to see if institution is inactive/closed according to NCES database
for cell in ws_uasys['U']:
    organization_name = str(cell.value)
    cell_prev = int(cell.row) - 1
    try:
        if cell_prev != 0 and organization_name != ws_uasys['U' + str(cell_prev)].value.upper():
            for look in ws_nces_grab['B']:
                nces_institution = str(look.value)
                if nces_institution.upper() == organization_name.upper():
                    institution_closed = ws_nces_grab['W' + str(look.row)].value
                    found_two = str(institution_closed).find('-2')
                    if found_two < 0:
                        ws_uasys['AI' + str(cell.row)].value = institution_closed
    except AttributeError:
        print("----------------------------------")
        print('NoneType for: ' + str(cell.value))
    except TypeError:
        print('NoneType')
    except:
        print('Unknown error')

# if INST_RECORD_SOURCE is blank then assign N/A
for cell in ws_uasys['AJ']:
    try:
        if cell.value is None:
            ws_uasys['AJ' + str(cell.row)].value = "N/A"
    except AttributeError:
        print('cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')
print('Done!')
# Change
wb_uasys.save(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Missouri Educational Institutions 2023-05-26.xlsx")

from openpyxl import workbook, load_workbook
import undetected_chromedriver as uc
import ssl
import bs4
import time
import random

raw_file = input("File location in explorer(.xlsx): ")
wrong_input = raw_file.find(".xlsx")
if wrong_input == -1:
    print("Be sure to add .xlsx to end of file location!")
    raw_file = input("File location is explorer(.xlsx)")
wb_uasys = load_workbook(raw_file)
wb_data_grab = load_workbook("AccreditationData.xlsx")
wb_nces_grab = load_workbook("Data_3-14-2023---623.xlsx")
sheet_name = input("Name of sheet in raw file: ")
ws_uasys = wb_uasys[sheet_name]
ws_data_grab = wb_data_grab["InstituteCampuses"]
ws_nces_grab = wb_nces_grab["Data_3-14-2023---623"]
abbrev = input("State Abbreviation of worksheet is needed: ")

# If CAMPUS_LOCATION_ID is blank then assign the cell AutoGen
for cell in ws_uasys['AK']:
    try:
        if cell.value is None:
            ws_uasys['AK' + str(cell.row)].value = "AutoGen"
    except AttributeError:
        print('Cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')
# Grabbing the missing cells in CAMP_OFFICIAL_INSTITUTION_NAME from PRIMARY_INSTITUTION_NAME
for cell in ws_uasys['AP']:
    if cell.value is None:
        try:
            CAMP_OFFICIAL_INSTITUTION_NAME = ws_uasys['U' + str(cell.row)].value
            ws_uasys['AP' + str(cell.row)].value = CAMP_OFFICIAL_INSTITUTION_NAME.upper()
        except AttributeError:
            print('NoneType object has no attribute upper')
# Get CAMP_OFFICIAL_INSTITUTION_NAME CAMP_OPED_ID and CAMP_IPED_ID from LocationName OpeId and IpedsUnitIds
for cell in ws_uasys['AP']:
    organization_name = str(cell.value)
    print("----------------------------------")
    print('Populating ' + organization_name + ' fields.....')
    for grab in ws_data_grab['D']:
        location_name = str(grab.value)
        if location_name.upper() == organization_name.upper():
            CAMP_DAPID = str(ws_data_grab['R' + str(grab.row)].value)
            CAMP_OPED_ID = str(ws_data_grab['S' + str(grab.row)].value)
            CAMP_IPED_ID = str(ws_data_grab['T' + str(grab.row)].value)
            ws_uasys['AL' + str(cell.row)].value = CAMP_DAPID
            ws_uasys['AM' + str(cell.row)].value = CAMP_OPED_ID
            ws_uasys['AN' + str(cell.row)].value = CAMP_IPED_ID
# Get CAMP_PO_BOX_LINE and CAMP_PhoneNumberFull from CAMP_OFFICIAL_INSTITUTION_NAME against LocationName fields
for cell in ws_uasys['AP']:
    organization_name = str(cell.value)
    for grab in ws_data_grab['D']:
        location_name = str(grab.value)
        if location_name.upper() == organization_name.upper():
            CAMP_PhoneNumberFull = str(ws_data_grab['I' + str(grab.row)].value)
            address_grab = str(ws_data_grab['H' + str(grab.row)].value)
            address_grab.split(', ')
            try:
                if len(address_grab.split(', ')) == 1:
                    address_grab = address_grab + ", N/A, N/A, N/A, N/A, N/A, N/A"
                if len(address_grab.split(', ')) == 2:
                    address_grab = address_grab + ", N/A, N/A, N/A, N/A, N/A"
                if len(address_grab.split(', ')) == 2:
                    address_grab = address_grab + ", N/A, N/A, N/A, N/A, N/A"
                if len(address_grab.split(', ')) == 3:
                    address_grab = address_grab + ", N/A, N/A, N/A, N/A"
                if len(address_grab.split(', ')) == 4:
                    address_grab = address_grab + ", N/A, N/A, N/A"

                GOV_ADDRESS_LINE_1, temp_LINE_2, temp_POBOX, temp_MUNI, temp_PCODE, temp1_Unknown, temp2_Unknown = address_grab.split(
                    ', ')

                if GOV_ADDRESS_LINE_1.startswith('P.O. Box'):
                    temp_PCODE = temp_POBOX
                    temp_POBOX = GOV_ADDRESS_LINE_1
                    GOV_ADDRESS_LINE_1 = 'N/A'

                if temp_POBOX.startswith('K'):
                    temp_POBOX = 'N/A'
                    temp_MUNI = temp_PCODE
                    temp_PCODE = temp1_Unknown

                if temp_PCODE.startswith('P.O BOX'):
                    temp_POBOX = temp_PCODE
                    temp_PCODE = 'NULL'

                if not temp_LINE_2.startswith('Suite'):
                    temp_MUNI = temp_LINE_2
                    temp_LINE_2 = 'N/A'
                    temp_PCODE = temp_POBOX
                    temp_POBOX = 'N/A'

                CAMP_ADDRESS_LINE_2 = temp_LINE_2.upper()
                CAMP_PO_BOX_LINE = temp_POBOX.strip('.')
                CAMP_MUNICIPALITY = temp_MUNI.upper()
                CAMP_POSTAL_CODE = temp_PCODE.strip(abbrev)

                ws_uasys['AT' + str(cell.row)].value = CAMP_ADDRESS_LINE_2
                ws_uasys['AU' + str(cell.row)].value = CAMP_PO_BOX_LINE
                ws_uasys['AV' + str(cell.row)].value = CAMP_MUNICIPALITY
                ws_uasys['AY' + str(cell.row)].value = CAMP_POSTAL_CODE
            except ValueError:
                ws_uasys['AT' + str(cell.row)].value = 'NULL'
                ws_uasys['AU' + str(cell.row)].value = 'NULL'
                ws_uasys['AV' + str(cell.row)].value = 'NULL'
                ws_uasys['AY' + str(cell.row)].value = 'NULL'
            except TypeError:
                ws_uasys['AT' + str(cell.row)].value = 'NULL'
                ws_uasys['AU' + str(cell.row)].value = 'NULL'
                ws_uasys['AV' + str(cell.row)].value = 'NULL'
                ws_uasys['AY' + str(cell.row)].value = 'NULL'
            except:
                ws_uasys['AT' + str(cell.row)].value = 'NULL'
                ws_uasys['AU' + str(cell.row)].value = 'NULL'
                ws_uasys['AV' + str(cell.row)].value = 'NULL'
                ws_uasys['AY' + str(cell.row)].value = 'NULL'

            ws_uasys['AZ' + str(cell.row)].value = CAMP_PhoneNumberFull
# Checking NCES for phonenumber if none is present
# Fix this
# for cell in ws_uasys['AQ']:
#     campus_name = str(cell.value)
#     for check in ws_uasys['AZ']:
#         if check.value is None:
#             print('No phone number from Accreditation Database : Searching')
#             for look in ws_nces_grab['B']:
#                 nces_institution = str(look.value)
#                 if nces_institution.upper() == campus_name.upper():
#                     print('Found a phone number number!')
#                     CAMP_PhoneNumberFull = str(ws_nces_grab['L' + str(look.row)].value)
#                     ws_uasys['AZ' + str(cell.row)].value = CAMP_PhoneNumberFull
# Get INST_ESTABLISHED_DATE for PRIMARY_INSTITUTION_NAME from Google search
# print('Looking up Institution established dates.........')
# for cell in ws_uasys['AP']:
#     PRIMARY_INSTITUTION_NAME = str(cell.value).upper()
#     if "BARBER" or "BEAUTY" or "HAIR" or "SALON" in PRIMARY_INSTITUTION_NAME is False:
#         try:
#             cell_prev = int(cell.row) - 1
#             if cell_prev != 0 and PRIMARY_INSTITUTION_NAME != ws_uasys['U' + str(cell_prev)].value.upper():
#                 print(PRIMARY_INSTITUTION_NAME + ' was founded:')
#                 if ws_uasys['AF' + str(cell.row)].value is None:
#                     ssl._create_default_https_context = ssl._create_unverified_context
#                     chrome_options = uc.ChromeOptions()
#
#                     url = 'https://google.com/search?q=' + '"' + str(PRIMARY_INSTITUTION_NAME) + '"' + ' / Founded'
#                     driver = uc.Chrome(options=chrome_options)
#                     driver.get(url)
#                     wait = random.randrange(1, 10)
#                     time.sleep(wait)
#                     request_result = driver.page_source
#                     driver.quit()
#                     web_data = bs4.BeautifulSoup(request_result, "html5lib")
#                     try:
#                         DATE = web_data.find('div', class_='Z0LcW t2b5Cf').text
#                         INST_ESTABLISHED_DATE = DATE
#                         print(INST_ESTABLISHED_DATE)
#                         ws_uasys['AF' + str(cell.row)].value = str(INST_ESTABLISHED_DATE) + '-01-01'
#                         wb_uasys.save(raw_file)
#                     except AttributeError:
#                         print("----------------------------------")
#                         print('NoneType for: ' + str(cell.value))
#                     except TypeError:
#                         print('NoneType')
#                     except:
#                         print('Unknown error')
#         except TypeError:
#             print('That was a merged or empty cell skipping......')
#         except AttributeError:
#             print('Cell is read only!')
#         except:
#             print('Unknown error')
# Check to see if campus is inactive/closed according to NCES database
for cell in ws_uasys['AP']:
    organization_name = str(cell.value)
    for look in ws_nces_grab['B']:
        nces_institution = str(look.value)
        if nces_institution.upper() == organization_name.upper():
            institution_closed = ws_nces_grab['W' + str(look.row)].value
            found_two = str(institution_closed).find('-2')
            if found_two < 0:
                ws_uasys['BD' + str(cell.row)].value = institution_closed
# If CAMPUS_RECORD_SOURCE is blank then assign the cell N/A
for cell in ws_uasys['BE']:
    try:
        if cell.value is None:
            ws_uasys['BE' + str(cell.row)].value = "N/A"
    except AttributeError:
        print('Cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')
print('Done!')
wb_uasys.save(raw_file)

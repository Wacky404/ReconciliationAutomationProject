import undetected_chromedriver as uc
import ssl
import re
import bs4
import time
import random


def governing_id(ws_uasys):
    # If GOVERNING_ORGANIZATION_ID is blank then assign the cell AutoGen
    for cell in ws_uasys['A']:
        try:
            if cell.value is None:
                ws_uasys['A' + str(cell.row)].value = "AutoGen"
        except AttributeError:
            print(cell + ' is read only!')


def governing_name(ws_uasys, ws_data_grab):
    # Get primary institution name and compare it against cells in additional sites location name,
    # if match: access Parent
    # Name Cell and return cell data to populate Governing Org name of same row as primary institution name
    for cell in ws_uasys['U']:
        institute_name = str(cell.value)
        print("----------------------------------")
        print(cell.value)

        for grab in ws_data_grab['D']:
            location_name = str(grab.value)

            if location_name.upper() == institute_name.upper():
                char = 'D'
                p_char = chr(ord(char) + 1)
                parent_name = str(ws_data_grab[p_char + str(grab.row)].value)

                if parent_name == "-":
                    parent_name = str(ws_data_grab['D' + str(grab.row)].value)

                ws_uasys['E' + str(cell.row)].value = parent_name
                print(ws_uasys['E' + str(cell.row)].value)
                print("----------------------------------")


def governing_edids(ws_uasys, ws_data_grab):
    # Get Governing_Organization_Name's DAPIP, OPE, and IPEDSID IDs from data_grab
    for cell in ws_uasys['E']:
        organization_name = str(cell.value)

        for grab in ws_data_grab['D']:
            location_name = str(grab.value)

            if location_name.upper() == organization_name.upper():
                GOV_DAPID = str(ws_data_grab['A' + str(grab.row)].value)
                GOV_OPEID = str(ws_data_grab['B' + str(grab.row)].value)
                GOV_IPEDID = str(ws_data_grab['C' + str(grab.row)].value)

                ws_uasys['B' + str(cell.row)].value = GOV_DAPID
                ws_uasys['C' + str(cell.row)].value = GOV_OPEID
                ws_uasys['D' + str(cell.row)].value = GOV_IPEDID


def governing_location(ws_uasys, ws_data_grab):
    # Get GOV address line 1, GOV_MUNICIPALITY, GOV postal code
    for cell in ws_uasys['E']:
        organization_name = str(cell.value)

        for grab in ws_data_grab['D']:
            location_name = str(grab.value)
            if location_name.upper() == organization_name.upper():
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

                    # if temp_PCODE.startswith('N/'):
                    #     temp_PCODE = temp_MUNI
                    #     temp_MUNI = temp_POBOX
                    #     temp_POBOX = "N/A"

                    if not temp_LINE_2.startswith('Suite'):
                        temp_MUNI = temp_LINE_2
                        temp_LINE_2 = 'N/A'
                        temp_PCODE = temp_POBOX
                        temp_POBOX = 'N/A'

                    GOV_ADDRESS_LINE_2 = temp_LINE_2.upper()
                    GOV_PO_BOX_LINE = temp_POBOX.strip('.')
                    GOV_MUNICIPALITY = temp_MUNI.upper()
                    GOV_POSTAL_CODE = temp_PCODE.strip(input('State Abbrev.: '))

                    ws_uasys['F' + str(cell.row)].value = GOV_ADDRESS_LINE_1
                    ws_uasys['G' + str(cell.row)].value = GOV_ADDRESS_LINE_2
                    ws_uasys['H' + str(cell.row)].value = GOV_PO_BOX_LINE
                    ws_uasys['I' + str(cell.row)].value = GOV_MUNICIPALITY
                    ws_uasys['L' + str(cell.row)].value = GOV_POSTAL_CODE
                except ValueError:
                    ws_uasys['F' + str(cell.row)].value = 'NULL'
                    ws_uasys['G' + str(cell.row)].value = 'NULL'
                    ws_uasys['H' + str(cell.row)].value = 'NULL'
                    ws_uasys['I' + str(cell.row)].value = 'NULL'
                    ws_uasys['L' + str(cell.row)].value = 'NULL'


def governing_region_short(ws_uasys):
    # If GOV_STATE_REGION_SHORT is blank then assign worksheet state
    for cell in ws_uasys['J']:
        try:
            if cell.value is None:
                # change state to workbook state
                ws_uasys['J' + str(cell.row)].value = input('State Abbrev.: ')
        except AttributeError:
            print('Cell is read only!')
    # If GOV_COUNTRY_CODE is blank then assign USA
    for cell in ws_uasys['K']:
        try:
            if cell.value is None:
                ws_uasys['K' + str(cell.row)].value = "USA"
        except AttributeError:
            print('Cell is read only!')


def governing_phone(ws_uasys, ws_data_grab):
    # Get GOV_PhoneNumberFull
    for cell in ws_uasys['E']:
        organization_name = str(cell.value)

        for grab in ws_data_grab['D']:
            location_name = str(grab.value)
            if location_name.upper() == organization_name.upper():
                phoneNumber_grab = str(ws_data_grab['I' + str(grab.row)].value)
                ws_uasys['M' + str(cell.row)].value = phoneNumber_grab


def governing_recordsource(ws_uasys):
    # If GOV_RECORD_SOURCE is blank then assign the cell N/A
    for cell in ws_uasys['P']:
        try:
            if cell.value is None:
                ws_uasys['P' + str(cell.row)].value = "N/A"
        except AttributeError:
            print(cell + ' is read only!')


def institution_id(ws_uasys):
    # If PRIMARY_ORGANIZATION_ID is blank then assign the cell AutoGen
    for cell in ws_uasys['Q']:
        try:
            if cell.value is None:
                ws_uasys['Q' + str(cell.row)].value = "AutoGen"
        except AttributeError:
            print(cell + ' is read only!')


def institution_po(ws_uasys, ws_data_grab):
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


def institution_country_code(ws_uasys):
    # If INST_COUNTRY_CODE is blank then assign USA
    for cell in ws_uasys['AA']:
        try:
            if cell.value is None:
                ws_uasys['AA' + str(cell.row)].value = "USA"
        except AttributeError:
            print('cell is read only!')


def institution_established_date(ws_uasys):
    # Get INST_ESTABLISHED_DATE for PRIMARY_INSTITUTION_NAME from Google search
    print('Looking up Institution established dates.........')
    for cell in ws_uasys['U']:
        PRIMARY_INSTITUTION_NAME = str(cell.value).upper()
        if "BARBER" or "BEAUTY" or "HAIR" or "SALON" in PRIMARY_INSTITUTION_NAME == False:
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
                            wb_uasys.save(
                                r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy TexasEducationalInstitutionsDatabase.xlsx")
                        except AttributeError:
                            print("----------------------------------")
                            print('NoneType for: ' + str(cell.value))
            except TypeError and AttributeError:
                print('That was a merged or empty cell skipping......')


def institution_recordsource(ws_uasys):
    # If INST_RECORD_SOURCE is blank then assign the cell N/A
    for cell in ws_uasys['AJ']:
        try:
            if cell.value is None:
                ws_uasys['AJ' + str(cell.row)].value = "N/A"
        except AttributeError:
            print(cell + ' is read only!')


def campus_id(ws_uasys):
    # If CAMPUS_LOCATION_ID is blank then assign the cell AutoGen
    for cell in ws_uasys['AK']:
        try:
            if cell.value is None:
                ws_uasys['AK' + str(cell.row)].value = "AutoGen"
        except AttributeError:
            print(cell + ' is read only!')


def campus_institution_edids(ws_uasys, ws_data_grab):
    # Get CAMP_OFFICIAL_INSTITUTION_NAME CAMP_OPED_ID and CAMP_IPED_ID from LocationName OpeId and IpedsUnitIds
    for cell in ws_uasys['AP']:
        organization_name = str(cell.value)
        print("----------------------------------")
        print('Populating ' + organization_name + ' fields.....')

        for grab in ws_data_grab['D']:
            location_name = str(grab.value)

            if location_name.upper() == organization_name.upper():
                # GOV_DAPID = str(ws_data_grab['A' + str(grab.row)].value)
                CAMP_OPED_ID = str(ws_data_grab['B' + str(grab.row)].value)
                CAMP_IPED_ID = str(ws_data_grab['C' + str(grab.row)].value)

                # ws_uasys['C' + str(cell.row)].value = GOV_DAPID
                ws_uasys['AM' + str(cell.row)].value = CAMP_OPED_ID
                ws_uasys['AN' + str(cell.row)].value = CAMP_IPED_ID


def campus_po_phonenumber(ws_uasys, ws_data_grab):
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
                    CAMP_POSTAL_CODE = temp_PCODE.strip('TX')

                    ws_uasys['AT' + str(cell.row)].value = CAMP_ADDRESS_LINE_2
                    ws_uasys['AU' + str(cell.row)].value = CAMP_PO_BOX_LINE
                    ws_uasys['AV' + str(cell.row)].value = CAMP_MUNICIPALITY
                    ws_uasys['AY' + str(cell.row)].value = CAMP_POSTAL_CODE
                except ValueError:
                    ws_uasys['AT' + str(cell.row)].value = 'NULL'
                    ws_uasys['AU' + str(cell.row)].value = 'NULL'
                    ws_uasys['AV' + str(cell.row)].value = 'NULL'
                    ws_uasys['AY' + str(cell.row)].value = 'NULL'

                ws_uasys['AZ' + str(cell.row)].value = CAMP_PhoneNumberFull

                if ws_uasys['AZ' + str(cell.row)].value is None:
                    print('No phone number from Accreditation Database : Searching')
                    for look in ws_nces_grab['B']:
                        nces_institution = str(look.value)
                        if nces_institution.upper() == organization_name.upper():
                            CAMP_PhoneNumberFull = str(ws_nces_grab['L' + str(grab.row)].value)
                            ws_uasys['AZ' + str(cell.row)].value = CAMP_PhoneNumberFull


def campus_recordsource(ws_uasys):
    # If CAMPUS_RECORD_SOURCE is blank then assign the cell N/A
    for cell in ws_uasys['BE']:
        try:
            if cell.value is None:
                ws_uasys['BE' + str(cell.row)].value = "N/A"
        except AttributeError:
            print(cell + ' is read only!')

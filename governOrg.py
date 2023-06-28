from openpyxl import workbook, load_workbook

# Change Add .xlsx
wb_uasys = load_workbook(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy California Educational Institutions 2023-06-20.xlsx")
wb_data_grab = load_workbook(r"C:\Users\Wayne Cole\Downloads\Work Stuff\AccreditationData.xlsx")
wb_nces_grab = load_workbook(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Data_3-14-2023---623.xlsx")
# Change
ws_uasys = wb_uasys["All California Institutions"]
ws_data_grab = wb_data_grab["InstituteCampuses"]
ws_nces_grab = wb_nces_grab["Data_3-14-2023---623"]
# If GOVERNING_ORGANIZATION_ID is blank then assign the cell AutoGen
for cell in ws_uasys['A']:
    try:
        if cell.value is None:
            ws_uasys['A' + str(cell.row)].value = "AutoGen"
    except AttributeError:
        print(cell + ' is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')
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
print("Populating associated fields.....hold on.....")
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

                if not temp_LINE_2.startswith('Suite'):
                    temp_MUNI = temp_LINE_2
                    temp_LINE_2 = 'N/A'
                    temp_PCODE = temp_POBOX
                    temp_POBOX = 'N/A'

                GOV_ADDRESS_LINE_2 = temp_LINE_2.upper()
                GOV_PO_BOX_LINE = temp_POBOX.strip('.')
                GOV_MUNICIPALITY = temp_MUNI.upper()
                # Change state abbreviation between states
                GOV_POSTAL_CODE = temp_PCODE.strip('CA')

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
            except:
                ws_uasys['F' + str(cell.row)].value = 'NULL'
                ws_uasys['G' + str(cell.row)].value = 'NULL'
                ws_uasys['H' + str(cell.row)].value = 'NULL'
                ws_uasys['I' + str(cell.row)].value = 'NULL'
                ws_uasys['L' + str(cell.row)].value = 'NULL'
# If GOV_STATE_REGION_SHORT is blank then assign worksheet state
for cell in ws_uasys['J']:
    try:
        if cell.value is None:
            # change state to workbook state
            ws_uasys['J' + str(cell.row)].value = "CA"
    except AttributeError:
        print('Cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')
# If GOV_COUNTRY_CODE is blank then assign USA
for cell in ws_uasys['K']:
    try:
        if cell.value is None:
            ws_uasys['K' + str(cell.row)].value = "USA"
    except AttributeError:
        print('Cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')
# Get GOV_PhoneNumberFull
for cell in ws_uasys['E']:
    organization_name = str(cell.value)

    for grab in ws_data_grab['D']:
        location_name = str(grab.value)
        if location_name.upper() == organization_name.upper():
            phoneNumber_grab = str(ws_data_grab['I' + str(grab.row)].value)
            ws_uasys['M' + str(cell.row)].value = phoneNumber_grab

            if ws_uasys['M' + str(cell.row)].value is None:
                print('No phone number from Accreditation Database : Searching')
                for look in ws_nces_grab['B']:
                    nces_institution = str(look.value)
                    if nces_institution.upper() == organization_name.upper():
                        phoneNumber_grab = str(ws_nces_grab['L' + str(grab.row)].value)
                        ws_uasys['M' + str(cell.row)].value = phoneNumber_grab
# Check to see if GOV_ORG is inactive/closed according to NCES database
for cell in ws_uasys['E']:
    organization_name = str(cell.value)
    for look in ws_nces_grab['B']:
        nces_institution = str(look.value)
        if nces_institution.upper() == organization_name.upper():
            institution_closed = ws_nces_grab['W' + str(look.row)].value
            found_two = str(institution_closed).find('-2')
            if found_two < 0:
                ws_uasys['O' + str(cell.row)].value = institution_closed
# If GOV_RECORD_SOURCE is blank then assign the cell N/A
for cell in ws_uasys['P']:
    try:
        if cell.value is None:
            ws_uasys['P' + str(cell.row)].value = "N/A"
    except AttributeError:
        print('Cell is read only!')
    except TypeError:
        print('Cell is read only!')
    except:
        print('Unknown error')

print('Done!')
# Change
wb_uasys.save(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy California Educational Institutions 2023-06-20.xlsx")

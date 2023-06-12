from openpyxl import workbook, load_workbook

wb_uasys = load_workbook(r"File Location")
wb_data_grab = load_workbook(r"File Location")
ws_uasys = wb_uasys["Name of Worksheet"]
ws_data_grab = wb_data_grab["Name of Worksheet"]
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

            # if CAMP_PhoneNumberFull == "":
            #     print('No phone number from Accreditation Database : Searching')
            #     # code for google search will go here
            # else:
            #     ws_uasys['BA' + str(cell.row)].value = CAMP_PhoneNumberFull
            #
            # ws_uasys['AV' + str(cell.row)].value = address_grab
print('Done!')
wb_uasys.save(r"File Location")

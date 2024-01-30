from openpyxl import load_workbook
from DataFile import DataFile
import openai
import time

# using gpt 4 is token expensive


class ReconcileAI:

    def __init__(self, raw_file, sheet_name, abbrev):
        self.raw_file = raw_file
        self.sheet_name = sheet_name
        self.abbrev = abbrev
        self.wb_uasys = load_workbook(raw_file)
        self.ws_uasys = self.wb_uasys[sheet_name]

    @classmethod
    def ai_institution(cls, wb_uasys, ws_uasys, raw_file):
        for cell in ws_uasys['U']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['Y' + str(cell.row)].value)
                    state = str(ws_uasys['Z' + str(cell.row)].value)
                    if institution_name != ws_uasys['U' + str(cell_prev)].value and ws_uasys[
                        'AF' + str(cell.row)].value is None:
                        API_KEY = open(r"C:\Users\Wayne\Work Stuff\Data Conversion\API Key.txt").read()
                        openai.api_key = API_KEY
                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "system", "content": "You are a data analyst reconciling missing data."},
                                {"role": "user",
                                 "content": "Don't include the question in your response, what is the date when"
                                            "Texas State University at San Marcos, TX founded?"},
                                {"role": "assistant", "content": "1899-01-01"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, what is the date when"
                                            "SAINT MARY'S COLLEGE OF CALIFORNIA at MORAGA, CA founded?"},
                                {"role": "assistant", "content": "1863-01-01"},
                                {"role": "user", "content": "If you can not find the date please respond with N/A."},
                                {"role": "assistant", "content": "N/A"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, what is the date when "
                                            + institution_name + ' at ' + municipality + ', ' + state + " founded?"}
                            ]
                        )
                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['AF' + str(cell.row)].value = str(reply_content)
                        else:
                            ws_uasys['AF' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['U' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['AF' + str(cell_prev)].value)
                        ws_uasys['AF' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                print(f"An exception of type {type(e).__name__} occurred in Insti. Details: {str(e)}")
                print('Moving on to the next location')

        for cell in ws_uasys['U']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['Y' + str(cell.row)].value)
                    state = str(ws_uasys['Z' + str(cell.row)].value)
                    if institution_name != ws_uasys['U' + str(cell_prev)].value and ws_uasys[
                        'AG' + str(cell.row)].value is None:
                        API_KEY = open(r"C:\Users\Wayne\Work Stuff\Data Conversion\API Key.txt").read()
                        openai.api_key = API_KEY
                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "system", "content": "You are a data analyst reconciling missing data."},
                                {"role": "user",
                                 "content": "Don't include the question in your response, When was this "
                                            "institution named Texas State University in San Marcos, TX?"},
                                {"role": "assistant", "content": "2013-01-01"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, When was this "
                                            "institution named SAINT MARY'S COLLEGE OF CALIFORNIA in MORAGA, CA?"},
                                {"role": "assistant", "content": "1863-01-01"},
                                {"role": "user", "content": "If you can not find the date please respond with N/A."},
                                {"role": "assistant", "content": "N/A"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, When was this "
                                            "institution named " + institution_name + ' in ' + municipality + ', '
                                            + state + "?"}
                            ]
                        )
                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['AG' + str(cell.row)].value = str(reply_content)
                        else:
                            ws_uasys['AG' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['U' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['AG' + str(cell_prev)].value)
                        ws_uasys['AG' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                print(f"An exception of type {type(e).__name__} occurred in Insti. Details: {str(e)}")
                print('Moving on to the next location')

    @classmethod
    def ai_campuslocation(cls, wb_uasys, ws_uasys, raw_file):
        for cell in ws_uasys['AP']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['AV' + str(cell.row)].value)
                    state = str(ws_uasys['AW' + str(cell.row)].value)
                    if institution_name != ws_uasys['AP' + str(cell_prev)].value and ws_uasys[
                        'BA' + str(cell.row)].value is None:
                        API_KEY = open(r"C:\Users\Wayne\Work Stuff\Data Conversion\API Key.txt").read()
                        openai.api_key = API_KEY
                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "system", "content": "You are a data analyst reconciling missing data."},
                                {"role": "user",
                                 "content": "Don't include the question in your response, what is the date when"
                                            "Texas State University at San Marcos, TX founded?"},
                                {"role": "assistant", "content": "1899-01-01"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, what is the date when"
                                            "SAINT MARY'S COLLEGE OF CALIFORNIA at MORAGA, CA founded?"},
                                {"role": "assistant", "content": "1863-01-01"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, what is the date when "
                                            + institution_name + ' at ' + municipality + ', ' + state + " founded?"}
                            ]
                        )
                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['BA' + str(cell.row)].value = str(reply_content)
                        else:
                            ws_uasys['BA' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['AP' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['BA' + str(cell_prev)].value)
                        ws_uasys['BA' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                print(f"An exception of type {type(e).__name__} occurred in Camp. Details: {str(e)}")
                print('Moving on to the next location')

        for cell in ws_uasys['AP']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['AV' + str(cell.row)].value)
                    state = str(ws_uasys['AW' + str(cell.row)].value)
                    if institution_name != ws_uasys['AP' + str(cell_prev)].value and ws_uasys[
                        'BB' + str(cell.row)].value is None:
                        API_KEY = open(r"C:\Users\Wayne\Work Stuff\Data Conversion\API Key.txt").read()
                        openai.api_key = API_KEY
                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "system", "content": "You are a data analyst reconciling missing data."},
                                {"role": "user",
                                 "content": "Don't include the question in your response, When was this "
                                            "campus named Texas State University in San Marcos, TX?"},
                                {"role": "assistant", "content": "2013-01-01"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, When was this "
                                            "campus named SAINT MARY'S COLLEGE OF CALIFORNIA in MORAGA, CA?"},
                                {"role": "assistant", "content": "1863-01-01"},
                                {"role": "user", "content": "If you can not find the date please respond with N/A."},
                                {"role": "assistant", "content": "N/A"},
                                {"role": "user",
                                 "content": "Don't include the question in your response, When was this "
                                            "campus named " + institution_name + ' in ' + municipality + ', '
                                            + state + "?"}
                            ]
                        )
                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['BB' + str(cell.row)].value = str(reply_content)
                        else:
                            ws_uasys['BB' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['AP' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['BB' + str(cell_prev)].value)
                        ws_uasys['BB' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                print(f"An exception of type {type(e).__name__} occurred in Camp. Details: {str(e)}")
                print('Moving on to the next location')

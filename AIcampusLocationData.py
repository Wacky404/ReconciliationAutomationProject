from openpyxl import load_workbook
import time
import openai


# def has_numbers(input_string):
#     return any(char.isdigit() for char in input_string)
#
#
# raw_file = input("File location in explorer(.xlsx): ")
# wrong_input = raw_file.find(".xlsx")
# if wrong_input == -1:
#     print("Be sure to add .xlsx to end of file location! ")
#     raw_file = input("File location is explorer(.xlsx)")
# wb_uasys = load_workbook(raw_file)
# sheet_name = input("Name of sheet in raw file: ")
# ws_uasys = wb_uasys[sheet_name]

# for cell in ws_uasys['AP']:
#     if cell.row >= 3:
#         cell_prev = int(cell.row) - 1
#         institution_name = str(cell.value)
#         municipality = str(ws_uasys['AV' + str(cell.row)].value)
#         state = str(ws_uasys['AW' + str(cell.row)].value)
#         if institution_name != ws_uasys['AP' + str(cell_prev)].value and ws_uasys['BA' + str(cell.row)].value is None:
#             API_KEY = open(r"C:\Users\Wayne Cole\Downloads\Work Stuff\API Key.txt").read()
#             openai.api_key = API_KEY
#             response = openai.ChatCompletion.create(
#                 model="gpt-3.5-turbo",
#                 messages=[
#                     {"role": "system", "content": "You are a data analyst reconciling missing data."},
#                     {"role": "user", "content": "Don't include the question in your response, what is the date when"
#                                                 "Texas State University at San Marcos, TX founded?"},
#                     {"role": "assistant", "content": "1899-01-01"},
#                     {"role": "user", "content": "Don't include the question in your response, what is the date when"
#                                                 "SAINT MARY'S COLLEGE OF CALIFORNIA at MORAGA, CA founded?"},
#                     {"role": "assistant", "content": "1863-01-01"},
#                     {"role": "user", "content": "If you can not find the date please respond with N/A."},
#                     {"role": "assistant", "content": "N/A"},
#                     {"role": "user", "content": "Don't include the question in your response, what is the date when "
#                                                 + institution_name + ' at ' + municipality + ', ' + state + " founded?"}
#                 ]
#             )
#             reply_content = response.choices[0].message.content
#             if has_numbers(reply_content):
#                 ws_uasys['BA' + str(cell.row)].value = str(reply_content)
#             else:
#                 ws_uasys['BA' + str(cell.row)].value = 'Manually Find'
#             wb_uasys.save(raw_file)
#             time.sleep(1)
# wb_uasys.save(raw_file)

# for cell in ws_uasys['AP']:
#     if cell.row >= 3:
#         cell_prev = int(cell.row) - 1
#         institution_name = str(cell.value)
#         municipality = str(ws_uasys['AV' + str(cell.row)].value)
#         state = str(ws_uasys['AW' + str(cell.row)].value)
#         if institution_name != ws_uasys['AP' + str(cell_prev)].value and ws_uasys['BB' + str(cell.row)].value is None:
#             API_KEY = open(r"C:\Users\Wayne Cole\Downloads\Work Stuff\API Key.txt").read()
#             openai.api_key = API_KEY
#             response = openai.ChatCompletion.create(
#                 model="gpt-3.5-turbo",
#                 messages=[
#                     {"role": "system", "content": "You are a data analyst reconciling missing data."},
#                     {"role": "user", "content": "Don't include the question in your response, When was this "
#                                                 "campus named Texas State University in San Marcos, TX?"},
#                     {"role": "assistant", "content": "2013-01-01"},
#                     {"role": "user", "content": "Don't include the question in your response, When was this "
#                                                 "campus named SAINT MARY'S COLLEGE OF CALIFORNIA in MORAGA, CA?"},
#                     {"role": "assistant", "content": "1863-01-01"},
#                     {"role": "user", "content": "If you can not find the date please respond with N/A."},
#                     {"role": "assistant", "content": "N/A"},
#                     {"role": "user", "content": "Don't include the question in your response, When was this "
#                                                 "campus named " + institution_name + ' in ' + municipality + ', '
#                                                 + state + "?"}
#                 ]
#             )
#             reply_content = response.choices[0].message.content
#             if has_numbers(reply_content):
#                 ws_uasys['BB' + str(cell.row)].value = str(reply_content)
#             else:
#                 ws_uasys['BB' + str(cell.row)].value = 'Manually Find'
#             wb_uasys.save(raw_file)
#             time.sleep(1)
# wb_uasys.save(raw_file)



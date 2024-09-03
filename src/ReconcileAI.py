# TODO: Fix credential error happening cause of DefaultAzureCredential()
from openpyxl import load_workbook
from src.DataFile import DataFile
from src.utils.log_util import logger
from openai import AzureOpenAI
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from dotenv import load_dotenv
import time
import os
import os.path as osp

load_dotenv()
ENDPOINT = os.getenv("ENDPOINT_URL")
DEPLOYMENT = os.getenv("DEPLOYMENT_NAME")
BEARER = os.getenv("TOKEN_URL")
API_VERSION = os.getenv("API_VERSION")
TOKEN_PROVIDER = get_bearer_token_provider(DefaultAzureCredential(), BEARER)


class ReconcileAI:

    def __init__(self, raw_file, sheet_name, abbrev):
        self.raw_file = raw_file
        base_file: str = osp.basename(self.raw_file)
        transf_f: str = osp.join(osp.expanduser(
            '~'), 'Documents', 'PipelineOutput', base_file)
        self.transf_file = transf_f if transf_f is not None else self.raw_file
        self.sheet_name = sheet_name
        self.abbrev = abbrev
        self.wb_uasys = load_workbook(raw_file)
        self.ws_uasys = self.wb_uasys[sheet_name]
        self.client = AzureOpenAI(
            azure_endpoint=ENDPOINT,
            azure_ad_token_provider=TOKEN_PROVIDER,
            api_version=API_VERSION,
        )

    def ai_institution(self, wb_uasys, ws_uasys, raw_file):
        for cell in ws_uasys['U']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['Y' + str(cell.row)].value)
                    state = str(ws_uasys['Z' + str(cell.row)].value)
                    if institution_name != ws_uasys['U' + str(cell_prev)].value and ws_uasys[
                        'AF' + str(cell.row)].value is None:
                        client = self.client
                        response = client.chat.completions.create(
                            model=DEPLOYMENT,
                            messages=[
                                {"role": "system",
                                 "content": "You are a data analyst reconciling missing data."},
                                {"role": "user",
                                 "content": "Don't include the question in your response and make sure the response "
                                            "is formated in Year-Month-Day, what is the date when "
                                            "Texas State University at San Marcos, TX founded?"},
                                {"role": "assistant", "content": "1899-01-01"},
                                {"role": "user",
                                 "content": "What is the date when "
                                            "SAINT MARY'S COLLEGE OF CALIFORNIA at MORAGA, CA founded?"},
                                {"role": "assistant", "content": "1863-01-01"},
                                {"role": "user",
                                 "content": "If you can not find the date when the place was founded respond with N/A."},
                                {"role": "assistant", "content": "N/A"},
                                {"role": "user",
                                 "content": "What is the date when "
                                            + institution_name + ' at ' + municipality + ', ' + state + " founded?"}
                            ],
                            temperature=0.7,
                            top_p=0.95,
                            frequency_penalty=0,
                            presence_penalty=0,
                            stop=None,
                            stream=False
                        )

                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['AF' + str(cell.row)
                                     ].value = str(reply_content)
                        else:
                            ws_uasys['AF' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['U' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['AF' + str(cell_prev)].value)
                        ws_uasys['AF' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                logger.exception(
                    f"An exception of type {type(e).__name__} occurred in Insti. Details: {str(e)}")
                logger.debug('Moving on to the next location')

        for cell in ws_uasys['U']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['Y' + str(cell.row)].value)
                    state = str(ws_uasys['Z' + str(cell.row)].value)
                    if institution_name != ws_uasys['U' + str(cell_prev)].value and ws_uasys[
                        'AG' + str(cell.row)].value is None:
                        client = self.client
                        response = client.chat.completions.create(
                            model=DEPLOYMENT,
                            messages=[
                                {"role": "system",
                                 "content": "You are a data analyst reconciling missing data."},
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
                            ],
                            temperature=0.7,
                            top_p=0.95,
                            frequency_penalty=0,
                            presence_penalty=0,
                            stop=None,
                            stream=False
                        )

                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['AG' + str(cell.row)
                                     ].value = str(reply_content)
                        else:
                            ws_uasys['AG' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['U' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['AG' + str(cell_prev)].value)
                        ws_uasys['AG' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                logger.exception(
                    f"An exception of type {type(e).__name__} occurred in Insti. Details: {str(e)}")
                logger.debug('Moving on to the next location')

    def ai_campuslocation(self, wb_uasys, ws_uasys, raw_file):
        for cell in ws_uasys['AP']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['AV' + str(cell.row)].value)
                    state = str(ws_uasys['AW' + str(cell.row)].value)
                    if institution_name != ws_uasys['AP' + str(cell_prev)].value and ws_uasys[
                        'BA' + str(cell.row)].value is None:
                        client = self.client
                        response = client.chat.completions.create(
                            model=DEPLOYMENT,
                            messages=[
                                {"role": "system",
                                 "content": "You are a data analyst reconciling missing data."},
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
                            ],
                            temperature=0.7,
                            top_p=0.95,
                            frequency_penalty=0,
                            presence_penalty=0,
                            stop=None,
                            stream=False
                        )
                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['BA' + str(cell.row)
                                     ].value = str(reply_content)
                        else:
                            ws_uasys['BA' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['AP' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['BA' + str(cell_prev)].value)
                        ws_uasys['BA' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                logger.exception(
                    f"An exception of type {type(e).__name__} occurred in Camp. Details: {str(e)}")
                logger.debug('Moving on to the next location')

        for cell in ws_uasys['AP']:
            try:
                if cell.row >= 3:
                    cell_prev = int(cell.row) - 1
                    institution_name = str(cell.value)
                    municipality = str(ws_uasys['AV' + str(cell.row)].value)
                    state = str(ws_uasys['AW' + str(cell.row)].value)
                    if institution_name != ws_uasys['AP' + str(cell_prev)].value and ws_uasys[
                        'BB' + str(cell.row)].value is None:
                        client = self.client
                        response = client.chat.completions.create(
                            model=DEPLOYMENT,
                            messages=[
                                {"role": "system",
                                 "content": "You are a data analyst reconciling missing data."},
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
                            ],
                            temperature=0.7,
                            top_p=0.95,
                            frequency_penalty=0,
                            presence_penalty=0,
                            stop=None,
                            stream=False
                        )
                        reply_content = response.choices[0].message.content
                        if DataFile.has_numbers(reply_content):
                            ws_uasys['BB' + str(cell.row)
                                     ].value = str(reply_content)
                        else:
                            ws_uasys['BB' + str(cell.row)].value = 'NULL'
                        wb_uasys.save(raw_file)
                        time.sleep(1)
                    elif institution_name == ws_uasys['AP' + str(cell_prev)].value:
                        last_entry = str(ws_uasys['BB' + str(cell_prev)].value)
                        ws_uasys['BB' + str(cell.row)].value = last_entry
                wb_uasys.save(raw_file)
            except Exception as e:
                logger.exception(
                    f"An exception of type {type(e).__name__} occurred in Camp. Details: {str(e)}")
                logger.debug('Moving on to the next location')

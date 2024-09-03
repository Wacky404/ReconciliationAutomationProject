from src.DataFile import DataFile as df
from src.ReconcileAI import ReconcileAI as ai
from src.NominatimIntegration import NominatimIntegration as nomi

from pathlib import Path
from src.utils.log_util import logger, setup_logging
import logging

import argparse
import os.path as osp
import os
import sys

import re
import time

os_home = osp.expanduser("~")
cwd = os.getcwd()
path_to_doc = osp.join(os_home, str('Documents'))
input_dir = osp.join(path_to_doc, str('Scheduled'))
output_dir = osp.join(path_to_doc, str('PipelineOutput'))
log_dir = osp.join(cwd, str('logs'))

parser = argparse.ArgumentParser(
    prog='DataPipeline',
    description="Data Reconciliation and Cleansing of Educational Institution data, using Excel."
                " Put your jobs that you want to be completed in a given pass in the input directory, Make sure "
                "that the file(s) is named <State>.xlsx. Once the workload is done the ouptut will be in the output "
                "directory."
                " The input directory will be cleared.",
)

parser.add_argument(
    '-c',
    '--configure',
    action='store_true',
    help='check/create directories that will be used in datapipeline',
)

parser.add_argument(
    '-l',
    '--log',
    action='store',
    default='INFO',  # NOTSET=0, DEBUG=10, INFO=20, WARNING=30, ERROR=40, CRITICAL=50
    help='DEBUG: Detailed information for diagnosing problems | '
         'INFO: Confirmation that things are working | '
         'WARNING: Indication that something unexpected happened. Program still running | '
         'ERROR: Not able to perform some function of the program | '
         'CRITICAL: Serious error, program may be unable to continue running',
)

parser.add_argument(
    '-t',
    '--task',
    action='store',
    default=int(2),
    # edit this help if you change user options
    help='Reconcile --> 1 | Reconcile+Cleanse --> 2 | Cleanse --> 3 | Reconcile+AI+Cleanse --> 4 | AI --> 5 | Test N '
         '--> 6',
)

# creates a NameSpace of arguments that were made
args = parser.parse_args()

numeric_loglevel = getattr(logging, str(args.log).upper())
if isinstance(numeric_loglevel, int):
    setup_logging(numeric_loglevel)
else:
    setup_logging()


def configure() -> None:
    """ Check/Create if directory exists that we will use to store excel workbooks to run on schedule """
    for directory in [input_dir, output_dir, log_dir]:
        try:
            os.makedirs(name=directory, exist_ok=False)
            print(f"Directory {directory} created")

        except FileExistsError as e:
            logger.exception(f"An exception of type {type(e).__name__} occurred. "
                             f"Details: This is okay, output will save in existing {directory}.")


if args.configure:
    configure()
    logger.info("Configuration completed.")
    sys.exit()

if osp.exists(input_dir):
    pathlist: list = list(Path(input_dir).glob('**/*.xlsx'))
    logger.info(f"Pathlist established: {pathlist}")
else:
    logger.critical(
        f"Dir: {input_dir} missing/denied; please check dir and/or run --configure if needed.")
    sys.exit()

if not osp.exists(output_dir):
    logger.critical(
        f"Dir: {output_dir} missing/denied; please check dir and/or run --configure if needed.")
    sys.exit()

if not osp.exists(log_dir):
    logger.critical(
        f"Dir: {log_dir} missing/denied; please check dir and/or run --configure if needed.")
    sys.exit()


def remove_fls(fls: list) -> None:
    """ Removes scheduled workload files from input directory """
    for f in fls:
        try:
            if osp.exists(f):
                os.remove(f)
            else:
                continue

        except (FileNotFoundError, OSError) as e:
            logger.exception(f"An exception of type {type(e).__name__} occurred. "
                             f"Details: {e} not found or is a directory")


print("Scheduled Workload:")
for path in pathlist:
    try:
        path = str(path)
        print(path)

    except Exception as e:
        logger.exception(f"An exception of type {type(e).__name__} occurred.")

filenames: list = list(re.sub(".xlsx$", "", osp.basename(file))
                       for file in pathlist)
zipped: list = list(zip(pathlist, filenames))

state: list = []
file_location: list = []
worksheet: list = []
abrev_state: list = []

amount: int = int(len(pathlist)) - 1
logger.info(f"Job amount(by index): {amount}")
i = 0
while i <= amount:
    try:
        input_state = str(zipped[i][1])
        state.append(input_state)

        input_file_location = str(zipped[i][0])
        file_location.append(input_file_location)

        input_worksheet = f"All {zipped[i][1]} Institutions"
        worksheet.append(input_worksheet)

        input_abrev_state = str(nomi.abbreviations[zipped[i][1]])
        abrev_state.append(input_abrev_state)

    except Exception as e:
        logger.exception(f"An exception of type {type(e).__name__} occurred. "
                         f"Details: Oops... check your filenames in Scheduled and make sure they are named correctly.")

    i += 1

logger.debug(f"{state}, {file_location}, {worksheet}, {abrev_state}")
logger.debug(f"This is task set: {type(args.task)} {args.task}")

user_task: int = int(args.task)
print('-------------------------------------------------------------------------------------')
if user_task == 1:
    logger.info(f'Data Operation: (1) has started...')

    for i in range(len(state)):
        start = time.perf_counter()
        state[i] = df(file_location[i], worksheet[i], abrev_state[i])

        logger.info(f'Reconcile Starting: {state[i].sheet_name}')

        state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.ws_data_grab,
                                       df.ws_nces_grab)
        state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, state[i].abbrev,
                                     df.ws_data_grab, df.ws_nces_grab)
        state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, state[i].abbrev,
                                          df.ws_data_grab, df.ws_nces_grab)
        state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file,
                                     df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

        print('Reconcile is done for ' + str(state[i].sheet_name) + '\n')
        logger.info(f'Reconcile Done: {state[i].sheet_name}')

        end = time.perf_counter()
        elasped = (end - start) * .000278

        logger.info(f'Reconcile Done: {state[i].sheet_name}, Time Elasped: {elasped:.2f} hrs')

    print('Removing jobs from Scheduled Dir...')
    remove_fls(file_location)
    logger.info('Workload is completed.')
    print('... Exiting')
    sys.exit()

elif user_task == 2:
    logger.info(f'Data Operation: (2) has started...')

    for i in range(len(state)):
        start = time.perf_counter()
        state[i] = df(file_location[i], worksheet[i], abrev_state[i])

        logger.info(f'Reconcile Starting: {state[i].sheet_name}')

        state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.ws_data_grab,
                                       df.ws_nces_grab)
        state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, state[i].abbrev,
                                     df.ws_data_grab, df.ws_nces_grab)
        state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, state[i].abbrev,
                                          df.ws_data_grab, df.ws_nces_grab)
        state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file,
                                     df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

        print('Reconcile is done for ' +
              str(state[i].sheet_name) + ' moving on to cleaning....\n')
        logger.info(f'Reconcile Done: {state[i].sheet_name}')

        state[i].clean_governing(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)
        state[i].clean_institution(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)
        state[i].clean_campuslocation(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)

        print('Clean is done for ' + str(state[i].sheet_name) + '\n')
        logger.info(f'Clean Done: {state[i].sheet_name}')

        end = time.perf_counter()
        elasped = (end - start) * .000278
        logger.info(f'Reconcile Done: {state[i].sheet_name}, Time Elasped: {elasped:.2f} hrs')

    print('Removing jobs from Scheduled Dir...')
    remove_fls(file_location)
    logger.info('Workload is completed.')
    print('... Exiting')
    sys.exit()

elif user_task == 3:
    logger.info(f'Data Operation: (3) has started...')

    for i in range(len(state)):
        start = time.perf_counter()
        state[i] = df(file_location[i], worksheet[i], abrev_state[i])

        print('Clean is starting for ' + str(state[i].sheet_name))

        state[i].clean_governing(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)
        state[i].clean_institution(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)
        state[i].clean_campuslocation(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)

        print('Clean is done for ' + str(state[i].sheet_name) + '\n')
        logger.info(f'Clean Done: {state[i].sheet_name}')

        end = time.perf_counter()
        elasped = (end - start) * .000278
        logger.info(f'Reconcile Done: {state[i].sheet_name}, Time Elasped: {elasped:.2f} hrs')

    print('Removing jobs from Scheduled Dir...')
    remove_fls(file_location)
    logger.info('Workload is completed.')
    print('... Exiting')
    sys.exit()

elif user_task == 4:
    logger.info(f'Data Operation: (4) has started...')

    for i in range(len(state)):
        start = time.perf_counter()
        state[i] = df(file_location[i], worksheet[i], abrev_state[i])

        logger.info(f'Reconcile Starting: {state[i].sheet_name}')

        state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.ws_data_grab,
                                       df.ws_nces_grab)
        state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, state[i].abbrev,
                                     df.ws_data_grab, df.ws_nces_grab)
        state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, state[i].abbrev,
                                          df.ws_data_grab, df.ws_nces_grab)
        state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file,
                                     df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

        print('Reconcile is done for ' +
              str(state[i].sheet_name) + ' moving on to AI....\n')
        logger.info(f'Reconcile Done: {state[i].sheet_name}')

        state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
        state_ai.ai_institution(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file)
        state_ai.ai_campuslocation(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file)

        print('AI is done for ' +
              str(state[i].sheet_name) + ' moving on to cleaning....\n')
        logger.info(f'AI Done: {state[i].sheet_name}')

        state[i].clean_governing(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)
        state[i].clean_institution(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)
        state[i].clean_campuslocation(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file, df.full_spellings)

        print('Clean is done for ' + str(state[i].sheet_name) + '\n')
        logger.info(f'Clean Done: {state[i].sheet_name}')

        end = time.perf_counter()
        elasped = (end - start) * .000278
        logger.info(f'Reconcile Done: {state[i].sheet_name}, Time Elasped: {elasped:.2f} hrs')

    print('Removing jobs from Scheduled Dir...')
    remove_fls(file_location)
    logger.info('Workload is completed.')
    print('... Exiting')
    sys.exit()

elif user_task == 5:
    logger.info(f'Data Operation: (5) has started...')

    for i in range(len(state)):
        start = time.perf_counter()
        state[i] = df(file_location[i], worksheet[i], abrev_state[i])

        print('AI is starting for ' + str(state[i].sheet_name))

        state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
        state_ai.ai_institution(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file)

        print('AI is done for ' +
              str(state[i].sheet_name) + ' institutions\n')

        state_ai.ai_campuslocation(
            state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file)

        print('AI is done for ' + str(state[i].sheet_name) + '\n')
        logger.info(f'AI Done: {state[i].sheet_name}')

        end = time.perf_counter()
        elasped = (end - start) * .000278
        logger.info(f'Reconcile Done: {state[i].sheet_name}, Time Elasped: {elasped:.2f} hrs')

    print('Removing jobs from Scheduled Dir...')
    remove_fls(file_location)
    logger.info('Workload is completed.')
    print('... Exiting')
    sys.exit()

elif user_task == 6:
    logger.info(f'Data Operation: (6) has started...')

    for i in range(len(state)):
        start = time.perf_counter()
        state[i] = df(file_location[i], worksheet[i], abrev_state[i])

        logger.info(f'Reconcile Starting: {state[i].sheet_name}')

        state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].transf_file,
                                     df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)
        end = time.perf_counter()
        elasped = (end - start) * .000278

        logger.info(f'Reconcile Nominatim Done: {state[i].sheet_name}, Time Elasped: {elasped:.2f} hrs')

    print('Removing jobs from Scheduled Dir...')
    remove_fls(file_location)
    logger.info('Workload is completed.')
    print('... Exiting')
    sys.exit()

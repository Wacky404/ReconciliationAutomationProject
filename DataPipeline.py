from DataFile import DataFile as df
from ReconcileAI import ReconcileAI as ai
from NominatimIntegration import NominatimIntegration as nomi
from pathlib import Path
import logging
import argparse
import os.path as osp
import os
import re

os_home = osp.expanduser("~")
path_to_doc = osp.join(os_home, str('Documents'))
input_dir = osp.join(path_to_doc, str('Scheduled'))
output_dir = osp.join(path_to_doc, str('PipelineOutput'))


def configure() -> None:
    """ Check/Create if directory exists that we will use to store excel workbooks to run on schedule """
    try:
        for directory in [input_dir, output_dir]:
            os.makedirs(name=directory, exist_ok=False)
            print(f"Directory {directory} created")

    except FileExistsError as e:
        print(f"An exception of type {type(e).__name__} occurred. "
              f"Details: This is okay, output will save in existing {directory}.")


def remove_fls(fls: list) -> None:
    """ Removes scheduled workload files from input directory """
    for f in fls:
        try:
            if osp.exists(f):
                os.remove(f)
            else:
                continue

        except (FileNotFoundError, OSError) as e:
            print(f"An exception of type {type(e).__name__} occurred. "
                  f"Details: {f} not found or is a directory")


parser = argparse.ArgumentParser(
    prog='DataPipeline',
    description='Data Reconciliation and Cleansing of Educational Institution data, using Excel.',
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
    help='DEBUG: Detailed information for diagnosing problems | '
    'INFO: Confirmation that things are working | '
    'WARNING: Indication that something unexpected happened. Program still running | '
    'ERROR: Not able to perform some function of the program | '
    'CRITICAL: Serious error, program may be unable to continue running',
)

# creates a NameSpace of arguments that were made
args = parser.parse_args()

if args.configure:
    configure()

if osp.exists(input_dir):
    pathlist: list = list(Path(input_dir).glob('**/*.xlsx'))
else:
    print(
        f"Dir: {input_dir} missing/denied; please check dir and/or run --configure if needed.")
    sys.exit()

if osp.exists(output_dir):
    continue
else:
    print(
        f"Dir: {output_dir} missing/denied; please check dir and/or run --configure if needed.")
    sys.exit()

print("Scheduled Workload:")
for path in pathlist:
    try:
        path = str(path)
        print(path)

    except Exception as e:
        print(f"An exception of type {type(e).__name__} occurred.")

filenames: list = list(re.sub(".xlsx$", "", osp.basename(file))
                       for file in pathlist)
zipped: list = list(zip(pathlist, filenames))

state: list = []
file_location: list = []
worksheet: list = []
abrev_state: list = []

amount: int = int(len(pathlist)) - 1
print('This is the length of pathlist: ', amount)
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
        print(f"An exception of type {type(e).__name__} occurred. "
              f"Details: Oops... check your filenames in Scheduled and make sure they are named correctly.")

    i += 1

print(state, file_location, worksheet, abrev_state, sep='\n')

go: bool = True
while go:
    print('-------------------------------------------------------------------------------------')
    user_choice: int = int(input('| Do you want to Reconcile --> 1 \n| Reconcile+Cleanse --> 2 \n| Cleanse --> 3 '
                                 '\n| Reconcile+AI+Cleanse --> 4 \n| AI --> 5 \n| Test N --> 6 \n'))
    if user_choice == 1:

        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab,
                                           df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                         df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                              df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file,
                                         df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

            print('Reconcile is done for ' + str(state[i].sheet_name) + '\n')

        print('Removing jobs from Scheduled Dir...')
        remove_fls(file_location)
        print('... Exiting')
        sys.exit()

    elif user_choice == 2:

        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab,
                                           df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                         df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                              df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file,
                                         df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

            print('Reconcile is done for ' +
                  str(state[i].sheet_name) + ' moving on to cleaning....\n')

            state[i].clean_governing(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(state[i].sheet_name) + '\n')

        print('Removing jobs from Scheduled Dir...')
        remove_fls(file_location)
        print('... Exiting')
        sys.exit()

    elif user_choice == 3:

        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            print('Clean is starting for ' + str(state[i].sheet_name))

            state[i].clean_governing(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(state[i].sheet_name) + '\n')

        print('Removing jobs from Scheduled Dir...')
        remove_fls(file_location)
        print('... Exiting')
        sys.exit()

    elif user_choice == 4:

        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab,
                                           df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                         df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                              df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file,
                                         df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

            print('Reconcile is done for ' +
                  str(state[i].sheet_name) + ' moving on to AI....\n')

            state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
            state_ai.ai_institution(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state_ai.ai_campuslocation(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' +
                  str(state[i].sheet_name) + ' moving on to cleaning....\n')

            state[i].clean_governing(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(state[i].sheet_name) + '\n')

        print('Removing jobs from Scheduled Dir...')
        remove_fls(file_location)
        print('... Exiting')
        sys.exit()

    elif user_choice == 5:

        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            print('AI is starting for ' + str(state[i].sheet_name))

            state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
            state_ai.ai_institution(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' +
                  str(state[i].sheet_name) + ' institutions\n')

            state_ai.ai_campuslocation(
                state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(state[i].sheet_name) + '\n')

        print('Removing jobs from Scheduled Dir...')
        remove_fls(file_location)
        print('... Exiting')
        sys.exit()

    elif user_choice == 6:

        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])
            state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file,
                                         df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)

        print('Removing jobs from Scheduled Dir...')
        remove_fls(file_location)
        print('... Exiting')
        sys.exit()

    else:
        print('You did not input any integer between 1 - 5, please try again\n')

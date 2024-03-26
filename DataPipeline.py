# TODO: Configure the script to check a folder for files and run them automatically, replacing user-input
from DataFile import DataFile as df
from ReconcileAI import ReconcileAI as ai
from pathlib import Path
import os.path as osp
import os

os_home = osp.expanduser("~")
path_to_doc = osp.join(os_home, str('Documents'))
input_dir = osp.join(path_to_doc, str('Scheduled'))
output_dir = osp.join(path_to_doc, str('PipelineOutput'))


def configure():
    """ Check if directory exists that we will use to store excel workbooks to run on schedule """
    try:
        for directory in [input_dir, output_dir]:
            os.makedirs(name=directory, exist_ok=False)
            print(f"Directory {directory} created")
    except FileExistsError as e:
        print(f"An exception of type {type(e).__name__} occurred. "
              f"Details: This is okay, output will save in existing directory.")


amount = int()
pathlist = Path(input_dir).glob('**/*.xlsx')
for path in pathlist:
    try:
        amount += 1
        path = str(path)
        print(path)
    except Exception as e:
        print(f"An exception of type {type(e).__name__} occurred.")

i = 0
while i < amount:
    try:
        state = []
        user_state = input('Name of state: ')
        state.append(user_state)

        file_location = []
        user_file_location = str(input('File location of workbook(.xlsx): '))
        file_location.append(user_file_location)

        worksheet = []
        user_worksheet = str(input('Exact worksheet name: '))
        worksheet.append(user_worksheet)

        abrev_state = []
        user_abrev_state = str(input('State abbreviation: '))
        abrev_state.append(user_abrev_state)
        i += 1
    except Exception as e:
        print(f"An exception of type {type(e).__name__} occurred. "
              f"Details: Oops... it looks like you didn't input an correct input, please try again")

go = True
while go:
    print('-------------------------------------------------------------------------------------')
    user_choice = int(input('| Do you want to Reconcile --> 1 \n| Reconcile+Cleanse --> 2 \n| Cleanse --> 3 '
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

            print('Reconcile is done for ' + str(state[i].sheet_name) + ' moving on to cleaning....\n')

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(state[i].sheet_name) + '\n')
    elif user_choice == 3:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            print('Clean is starting for ' + str(state[i].sheet_name))

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(state[i].sheet_name) + '\n')
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

            print('Reconcile is done for ' + str(state[i].sheet_name) + ' moving on to AI....\n')

            state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
            state_ai.ai_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state_ai.ai_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(state[i].sheet_name) + ' moving on to cleaning....\n')

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(state[i].sheet_name) + '\n')
    elif user_choice == 5:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            print('AI is starting for ' + str(state[i].sheet_name))

            state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
            state_ai.ai_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(state[i].sheet_name) + ' institutions\n')

            state_ai.ai_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(state[i].sheet_name) + '\n')
    elif user_choice == 6:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])
            state[i].reconcile_nominatim(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file,
                                         df.null_values, df.gov_field_names, df.insti_field_names, df.camp_field_names)
    else:
        print('You did not input any integer between 1 - 5, please try again\n')

    exit = int(input('Do you want to exit: Yes --> 0 | No --> 1 '))
    if exit == 1:
        continue
    else:
        break

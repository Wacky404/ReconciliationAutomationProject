# googleintegration is going to need to be changed same as ReconcileAI file
from DataFile import DataFile as df
from ReconcileAI import ReconcileAI as ai

try:
    amount = int(input('How many states are you reconciling/cleansing?(integer): '))
except Exception as e:
    print(f"An exception of type {type(e).__name__} occurred. Details: {str(e)}")
    amount = int(input("Oops... it looks like you didn't input an integer, please try again(integer)"))

i = 0
while i < amount:
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
go = True
while go:
    user_choice = int(input('Do you want to Reconcile --> 1 | Reconcile+Cleanse --> 2 | Cleanse --> 3 '
                            '| Reconcile+AI+Cleanse --> 4 | AI --> 5 '))
    if user_choice == 1:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab,
                                           df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                         df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                              df.ws_data_grab, df.ws_nces_grab)
            #state[i].reconcile_google(state[i].wb_uasys, state[i].ws_uasys, df.null_values)

            print('Reconcile is done for ' + str(*state[i]) + '\n')
    elif user_choice == 2:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab,
                                           df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                         df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                              df.ws_data_grab, df.ws_nces_grab)
            #state[i].reconcile_google(state[i].wb_uasys, state[i].ws_uasys, df.null_values)

            print('Reconcile is done for ' + str(*state[i]) + ' moving on to cleaning....\n')

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(*state[i]) + '\n')
    elif user_choice == 3:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(*state[i]) + '\n')
    elif user_choice == 4:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab,
                                           df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                         df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev,
                                              df.ws_data_grab, df.ws_nces_grab)
            #state[i].reconcile_google(state[i].wb_uasys, state[i].ws_uasys, df.null_values)

            print('Reconcile is done for ' + str(*state[i]) + ' moving on to AI....\n')

            state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
            state_ai.ai_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state_ai.ai_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(*state[i]) + ' moving on to cleaning....\n')

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.full_spellings)

            print('Clean is done for ' + str(*state[i]) + '\n')
    elif user_choice == 5:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state_ai = ai(file_location[i], worksheet[i], abrev_state[i])
            state_ai.ai_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(*state[i]) + ' institutions\n')

            state_ai.ai_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)

            print('AI is done for ' + str(*state[i]) + '\n')
    else:
        print('You did not input any integer between 1 - 5, please try again\n')

    exit = int(input('Do you want to exit: Yes --> 0 | No --> 1 '))
    if exit == 1:
        continue
    else:
        break
